#!/usr/bin/env python3
"""Check the project's native, chapter-local caption field contract.

This is a structural check, not an Office field evaluator.  It accepts both
fldSimple and complete begin/instruction/separate/result/end representations.
Neither heading text nor style display names are used as SEQ arguments.
"""
from __future__ import annotations

import argparse
from dataclasses import dataclass, field
import hashlib
import json
import os
import shutil
import tempfile
from pathlib import Path
import re
import zipfile
from lxml import etree as E

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}
Q = lambda name: '{' + W + '}' + name
FALSE = {'0', 'false', 'off'}
LABELS = {'figure': '图', 'table': '表'}
PLACEMENT = {'figure': 'after', 'table': 'before'}
ALIASES = {'图': 'figure', '表': 'table', 'figure': 'figure', 'table': 'table'}
ERROR_TEXT = re.compile(
    r'(?:错误[!！]\s*(?:未定义|未找到|文档中没有|找不到)|'
    r'Error!\s*(?:No text|Reference source|Bookmark|Not a valid|Unknown|Syntax))', re.I)
TOKEN = re.compile(r'"(?:[^"\r\n]|"")*"|\\[a-zA-Z*#@]|[^\s\\"]+')
CAPTION_LABEL = re.compile(r'^\s*(图|表|Figure|Table)\s*(?=\d|错误[!！]|Error!)', re.I)


class CaptionFieldError(ValueError):
    def __init__(self, code: str, message: str, **details):
        super().__init__(message)
        self.code, self.details = code, details

    def as_issue(self) -> dict:
        return {'severity': 'error', 'code': self.code, 'message': str(self), **self.details}


def parse(data: bytes):
    return E.fromstring(data, E.XMLParser(resolve_entities=False, no_network=True))


def value(el, name='val', default=None):
    return default if el is None else el.get(Q(name), default)


def digest(path) -> str:
    return 'sha256:' + hashlib.sha256(Path(path).read_bytes()).hexdigest()


def own_elements(p):
    """Do not count a textbox's nested paragraph as part of its owner's field."""
    def walk(el):
        for child in el:
            if child.tag in {Q('p'), Q('del'), Q('moveFrom')}:
                continue
            yield child
            yield from walk(child)
    yield from walk(p)


def text(p) -> str:
    return ''.join((n.text or '') for n in own_elements(p) if n.tag == Q('t'))


@dataclass
class Field:
    instruction: str = ''
    cached: str = ''
    complete: bool = False
    locked: bool = False
    nested: bool = False
    separated: bool = False
    errors: list[str] = field(default_factory=list)
    simple_node: object = None
    instruction_nodes: list = field(default_factory=list)


def fields(p) -> tuple[list[Field], list[str]]:
    """Read native fields without pretending to calculate their results."""
    result: list[Field] = []
    stack: list[Field] = []
    malformed: list[str] = []

    def walk(el):
        for n in el:
            if n.tag in {Q('p'), Q('del'), Q('moveFrom')}:
                continue
            if n.tag == Q('fldSimple'):
                f = Field(instruction=value(n, 'instr', ''), cached=text(n),
                          complete=True, locked=value(n, 'fldLock', '0').lower() not in FALSE,
                          nested=bool(stack) or bool(n.xpath('.//w:fldSimple|.//w:fldChar', namespaces=NS)))
                f.simple_node = n
                result.append(f)
                for outer in stack:
                    outer.nested = True
                    if outer.separated:
                        outer.cached += f.cached
                continue
            if n.tag == Q('fldChar'):
                typ = value(n, 'fldCharType')
                if typ == 'begin':
                    f = Field(locked=value(n, 'fldLock', '0').lower() not in FALSE, nested=bool(stack))
                    for outer in stack:
                        outer.nested = True
                    result.append(f)
                    stack.append(f)
                elif typ == 'separate':
                    if not stack:
                        malformed.append('孤立的域结果分隔标记')
                    elif stack[-1].separated:
                        malformed.append('同一域含重复结果分隔标记')
                    else:
                        stack[-1].separated = True
                elif typ == 'end':
                    if not stack:
                        malformed.append('孤立的域结束标记')
                    else:
                        stack.pop().complete = True
                else:
                    malformed.append('未知的域边界类型')
            elif n.tag == Q('instrText'):
                if stack and not stack[-1].separated:
                    stack[-1].instruction += n.text or ''
                    stack[-1].instruction_nodes.append(n)
                else:
                    malformed.append('域指令不在有效边界内')
            elif n.tag == Q('t'):
                for f in stack:
                    if f.separated:
                        f.cached += n.text or ''
            walk(n)
    walk(p)
    if stack:
        malformed.append('域开始标记没有对应结束标记')
    return result, malformed


def tokens(instruction: str) -> list[str]:
    out, end = [], 0
    for m in TOKEN.finditer(instruction):
        if instruction[end:m.start()].strip():
            raise CaptionFieldError('caption-seq-syntax', '题注域含无法解析的引号或反斜杠。')
        out.append(m.group()); end = m.end()
    if instruction[end:].strip():
        raise CaptionFieldError('caption-seq-syntax', '题注域指令未闭合。')
    return out


def command(instruction: str) -> str:
    return instruction.strip().split(None, 1)[0].upper() if instruction.strip() else ''


def sequence_kind(instruction: str):
    m = re.match(r'^\s*SEQ\s+("[^"]+"|[^\s\\]+)', instruction, re.I)
    return ALIASES.get(m[1].strip('"').lower()) if m else None


def check_instruction(instruction: str, kind: str, *, reset=False) -> None:
    ts = tokens(instruction)
    if len(ts) < 2 or ts[0].upper() != 'SEQ' or ts[1].strip('"') != LABELS[kind]:
        raise CaptionFieldError('caption-seq-identity', '图、表必须使用各自固定的 SEQ 序列，不得使用标题文字或样式名作序列名。')
    switches: dict[str, str | bool | list[str]] = {}
    i = 2
    while i < len(ts):
        switch = ts[i].lower(); i += 1
        if switch not in {'\\*', '\\s', '\\r', '\\h', '\\n'}:
            raise CaptionFieldError('caption-seq-dependency', '题注域含书签参数、不支持的开关或多余参数；不得依赖标题引用。', token=switch)
        if switch in switches and switch != '\\*':
            raise CaptionFieldError('caption-seq-duplicate-switch', '题注域含重复开关。', switch=switch)
        if switch in {'\\*', '\\s', '\\r'}:
            if i >= len(ts) or ts[i].startswith('\\'):
                raise CaptionFieldError('caption-seq-argument', '题注域缺少开关参数。', switch=switch)
            argument = ts[i]; i += 1
            if switch == '\\*':
                switches.setdefault(switch, []).append(argument.upper())
            else:
                switches[switch] = argument
        else:
            switches[switch] = True
    if reset:
        if switches != {'\\r': '0', '\\h': True}:
            raise CaptionFieldError('caption-reset-invalid', '章级隐藏重置只能使用 SEQ 图/表 \\r 0 \\h。')
    else:
        if switches.get('\\s') != '1' or any(k in switches for k in ('\\r', '\\h')):
            raise CaptionFieldError('caption-seq-reset', '可见题注应按一级章节重启；不得逐条固定编号、隐藏或把样式名当重启层级。')
        formats = switches.get('\\*', [])
        if any(x not in {'ARABIC', 'MERGEFORMAT', 'CHARFORMAT'} for x in formats) or len(formats) != len(set(formats)):
            raise CaptionFieldError('caption-seq-format', '题注序号必须使用阿拉伯数字。')


def style_chain(styles: dict, sid):
    seen = set()
    while sid:
        if sid in seen:
            raise CaptionFieldError('caption-style-cycle', '样式继承循环，无法核对题注及章节层级。')
        seen.add(sid)
        style = styles.get(sid)
        if style is None:
            break
        yield style
        sid = value(style.find(Q('basedOn')))


def paragraph_info(p, styles: dict) -> tuple[list[str], int | None]:
    sid = value(p.find('w:pPr/w:pStyle', NS))
    chain = list(style_chain(styles, sid))
    names = [value(s.find(Q('name')), default='') for s in chain]
    outline = p.find('w:pPr/w:outlineLvl', NS)
    if outline is None:
        outline = next((s.find('w:pPr/w:outlineLvl', NS) for s in chain
                        if s.find('w:pPr/w:outlineLvl', NS) is not None), None)
    lv = None
    if outline is not None:
        try:
            lv = int(value(outline))
        except (ValueError, TypeError):
            pass
    if lv is None:
        for name in names:
            m = re.fullmatch(r'(?:heading|标题)\s*([1-9])', name, re.I)
            if m:
                lv = int(m[1]) - 1; break
    if p.xpath('ancestor::w:tc|ancestor::w:txbxContent', namespaces=NS):
        lv = None
    return names, lv


def _near_object(p) -> bool:
    for forward in (False, True):
        n = p.getnext() if forward else p.getprevious()
        for _ in range(3):
            if n is None:
                break
            if n.tag == Q('tbl') or n.xpath('.//w:drawing|.//w:pict', namespaces=NS):
                return True
            if n.tag != Q('p') or text(n).strip():
                break
            n = n.getnext() if forward else n.getprevious()
    return False


def audit(path, *, check_cache=True) -> dict:
    """Supplement the core's object/placement audit; never suppress core issues."""
    path = Path(path)
    with zipfile.ZipFile(path) as z:
        root = parse(z.read('word/document.xml'))
        styles_root = parse(z.read('word/styles.xml'))
    styles = {value(s, 'styleId'): s for s in styles_root.findall(Q('style'))}
    issues, records, chapter, counters = [], [], 0, {}

    def bad(code, message, **detail):
        issues.append({'severity': 'error', 'code': code, 'message': message, **detail})

    for index, p in enumerate(root.iter(Q('p')), 1):
        if p.xpath('ancestor::w:del|ancestor::w:moveFrom', namespaces=NS):
            continue
        names, level = paragraph_info(p, styles)
        if any(re.match(r'^(?:TOC|目录)', name, re.I) for name in names):
            continue
        txt = text(p).strip()
        if level == 0 and txt:
            chapter += 1
        fs, malformed = fields(p)
        seqs = [f for f in fs if sequence_kind(f.instruction)]
        caption_style = any(re.search(r'caption|题注|图题|表题', n, re.I) for n in names)
        is_caption = bool(caption_style and txt) or bool(CAPTION_LABEL.match(txt) and _near_object(p))
        if not is_caption and not seqs:
            continue
        if malformed:
            bad('caption-field-boundaries', '题注/重置域边界损坏，不能用缓存代替可更新域。', paragraph=index, details=malformed)
        for f in fs:
            if ERROR_TEXT.search(f.cached):
                bad('caption-field-result-error', '题注域结果中仍有引用或样式错误。', paragraph=index, result=f.cached)
        if is_caption:
            kinds = {sequence_kind(f.instruction) for f in seqs}
            if len(seqs) != 1 or len(kinds) != 1:
                bad('caption-native-seq-required', '题注必须有且只有一个真实 SEQ 图/表域。', paragraph=index)
                continue
            f = seqs[0]; kind = next(iter(kinds))
            for other in fs:
                if other is not f and command(other.instruction) in {'STYLEREF', 'REF', 'PAGEREF', 'SEQ'}:
                    bad('caption-heading-dependency', '题注中不得拼接标题样式引用、书签引用或第二个编号域。', paragraph=index)
            if not f.complete or f.nested or f.locked:
                bad('caption-field-not-updateable', '题注序号域未闭合、嵌套或锁定。', paragraph=index)
            try:
                check_instruction(f.instruction, kind)
            except CaptionFieldError as exc:
                issues.append({**exc.as_issue(), 'paragraph': index})
            counters[(chapter, kind)] = counters.get((chapter, kind), 0) + 1
            number = counters[(chapter, kind)]
            if check_cache and f.cached.strip() != str(number):
                bad('caption-refresh-sequence', '题注域结果不符合当前章节实际对象顺序；须真实更新域后复核。',
                    paragraph=index, expected=number, actual=f.cached)
            if check_cache and not re.match(r'^' + LABELS[kind] + str(number) + r'\s+\S', txt):
                bad('caption-display-prefix', '图表题注应显示图1/表1及题名，不拼章节前缀。', paragraph=index)
            records.append({'paragraph': index, 'chapter': chapter, 'kind': kind,
                            'number': f.cached.strip(), 'expected_number': number, 'instruction': f.instruction.strip()})
        else:
            # Core-generated invisible resets may only live on the chapter heading.
            used = set()
            for f in seqs:
                kind = sequence_kind(f.instruction)
                try:
                    check_instruction(f.instruction, kind, reset=True)
                except CaptionFieldError as exc:
                    issues.append({**exc.as_issue(), 'paragraph': index})
                if level != 0 or kind in used or not f.complete or f.nested or f.locked:
                    bad('caption-reset-location', '隐藏重置必须位于一级标题且每种序列至多一个；不能位于普通正文或首张图表。', paragraph=index)
                used.add(kind)
    return {'status': 'failed' if issues else 'passed', 'input': str(path), 'input_sha256': digest(path),
            'issues': issues, 'caption_count': len(records), 'captions': records,
            'position_policy': PLACEMENT, 'office_refresh_proven': False,
            'note': '结构检查不等同于实际 Office 更新域；更新后仍须检查序号及错误结果。'}


def require_valid(path, *, check_cache=True) -> dict:
    result = audit(path, check_cache=check_cache)
    if result['status'] != 'passed':
        raise CaptionFieldError('caption-field-contract-failed', '图表自动编号未满足稳定刷新基准。', issues=result['issues'])
    return result


def _literal_text(p) -> str:
    """Stable paragraph anchor, excluding field results and run segmentation."""
    out, depth = [], 0
    def walk(el):
        nonlocal depth
        for n in el:
            if n.tag in {Q('p'), Q('del'), Q('moveFrom'), Q('fldSimple')}:
                continue
            if n.tag == Q('fldChar'):
                typ = value(n, 'fldCharType')
                if typ == 'begin': depth += 1
                elif typ == 'end': depth -= 1
            elif n.tag == Q('t') and depth == 0:
                out.append(n.text or '')
            elif n.tag == Q('tab') and depth == 0:
                out.append('\t')
            elif n.tag in {Q('br'), Q('cr')} and depth == 0:
                out.append('\n')
            walk(n)
    walk(p)
    return ''.join(out)


def _native_entries(root):
    entries = []
    for p in root.iter(Q('p')):
        if p.xpath('ancestor::w:del|ancestor::w:moveFrom', namespaces=NS):
            continue
        fs, malformed = fields(p)
        monitored = [f for f in fs if sequence_kind(f.instruction)]
        if monitored and (malformed or any(not f.complete or f.nested or f.locked for f in monitored)):
            raise CaptionFieldError('caption-roundtrip-boundaries', 'Office 输出的域边界已损坏，不能按猜测恢复。')
        for index, f in enumerate(monitored):
            entries.append((_literal_text(p), index, sequence_kind(f.instruction), f))
    return entries


def preserve_native_instructions(source, office_output, out) -> dict:
    """Restore only source-validated instructions after Office serialization.

    Some Office exporters discard SEQ switches while leaving a correct first
    cached value. Copying those files would make the next F9 unsafe. Match every
    managed field by its paragraph's literal text, field order and sequence kind;
    restore the *validated source instruction only*. Keep the engine's result
    text unchanged. Missing/reordered objects or bad recalculations fail closed.
    This is a serialization-preservation part of refresh, not synthetic F9.
    """
    source, office_output, out = map(Path, (source, office_output, out))
    if out.resolve() in {source.resolve(), office_output.resolve()}:
        raise CaptionFieldError('caption-roundtrip-output-collision', '原生域保真输出不能覆盖输入或原始引擎结果。')
    require_valid(source, check_cache=False)
    with zipfile.ZipFile(source) as z:
        original = parse(z.read('word/document.xml'))
    with zipfile.ZipFile(office_output) as z:
        final = parse(z.read('word/document.xml'))
    before, after = _native_entries(original), _native_entries(final)
    if len(before) != len(after):
        raise CaptionFieldError('caption-roundtrip-count', 'Office 保存前后图表域数量不同，不能按旧位置恢复。')
    changes = []
    cached_before = [entry[3].cached for entry in after]
    for index, (old, new) in enumerate(zip(before, after), 1):
        if old[:3] != new[:3]:
            raise CaptionFieldError('caption-roundtrip-anchor', 'Office 保存改变了题注/标题的对应关系，不能按猜测恢复指令。', field=index)
        expected, actual = old[3], new[3]
        if expected.instruction.strip() == actual.instruction.strip():
            continue
        if actual.simple_node is not None:
            actual.simple_node.set(Q('instr'), expected.instruction)
        elif actual.instruction_nodes:
            actual.instruction_nodes[0].text = expected.instruction
            actual.instruction_nodes[0].set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            for n in actual.instruction_nodes[1:]: n.text = '' 
        else:
            raise CaptionFieldError('caption-roundtrip-instruction', 'Office 输出缺少可定位的原生域指令。', field=index)
        changes.append({'field': index, 'kind': old[2], 'before': actual.instruction.strip(),
                        'restored': expected.instruction.strip(), 'cached_result': actual.cached})
    if cached_before != [entry[3].cached for entry in _native_entries(final)]:
        raise CaptionFieldError('caption-roundtrip-cache-mutated', '恢复域指令不得改写引擎实际计算的结果。')
    out.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(dir=out.parent) as td:
        candidate = Path(td) / 'preserved.docx'
        if changes:
            xml = E.tostring(final, encoding='UTF-8', xml_declaration=True, standalone=True)
            with zipfile.ZipFile(office_output) as src, zipfile.ZipFile(candidate, 'w') as dst:
                for info in src.infolist():
                    dst.writestr(info, xml if info.filename == 'word/document.xml' else src.read(info.filename))
        else:
            shutil.copyfile(office_output, candidate)
        checked = require_valid(candidate)
        os.replace(candidate, out)
    return {'status': 'passed', 'source_sha256': digest(source), 'office_output_sha256': digest(office_output),
            'output_sha256': digest(out), 'instruction_restore_count': len(changes), 'changes': changes,
            'engine_cached_results_unchanged': True, 'caption_count': checked['caption_count'],
            'note': '仅恢复已验证的源域指令；序号缓存来自真实 Office 结果，未自行计算或写入假缓存。'}


def main():
    ap = argparse.ArgumentParser(description='核对固定图表序列及标题无关的域指令')
    ap.add_argument('input', type=Path)
    ap.add_argument('--before-refresh', action='store_true', help='只核对域结构，允许待更新的旧序号缓存')
    ap.add_argument('--json-out', type=Path)
    args = ap.parse_args()
    try:
        result = audit(args.input, check_cache=not args.before_refresh)
    except Exception as exc:
        result = {'status': 'failed', 'issues': [exc.as_issue() if hasattr(exc, 'as_issue') else
                  {'severity': 'error', 'code': 'caption-field-read-error', 'message': str(exc)}]}
    payload = json.dumps(result, ensure_ascii=False, indent=2)
    if args.json_out:
        if args.json_out.resolve() == args.input.resolve():
            raise SystemExit('报告不能覆盖输入文档')
        args.json_out.parent.mkdir(parents=True, exist_ok=True)
        args.json_out.write_text(payload + '\n', encoding='utf-8')
    print(payload)
    return 0 if result['status'] == 'passed' else 2


if __name__ == '__main__':
    raise SystemExit(main())
