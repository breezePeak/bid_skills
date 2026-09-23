#!/usr/bin/env python3
"""Template-authoritative heading numbers and chapter-local Word captions.

No image OCR or image-quality claims are made here. Image content is checked by
figure_review.py using current-image + final-page review evidence. Existing
footnote conversion remains in numbering_core.py; see the CLI adapters.
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
import os
import posixpath
import re
import tempfile
import zipfile
from collections import defaultdict
from pathlib import Path
from typing import Any

from lxml import etree as E

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
V = 'urn:schemas-microsoft-com:vml'
NS = {'w': W, 'r': R, 'a': A, 'wp': WP, 'v': V}
Q = lambda n: '{' + W + '}' + n
XMLSPACE = '{http://www.w3.org/XML/1998/namespace}space'
FALSE = {'0', 'false', 'off'}
CN = '零〇一二三四五六七八九十百千万两'
CAP = re.compile(r'^\s*(图|表|Figure|Table)\s*(\d+(?:[.．\-－]\d+)*)(?:[.．、])?\s*', re.I)
REF = re.compile(r'(图|表|Figure|Table)\s*(\d+(?:[.．\-－]\d+)*)', re.I)
ERROR_TEXT = re.compile(r'错误[！!]\s*(?:未定义样式|未找到引用源|未定义书签)|Error!\s*(?:No text of specified style|Reference source not found|Bookmark not defined)', re.I)
SEQUENCES = {'figure': '图', 'table': '表'}
LABELS = {'figure': '图', 'table': '表'}


class PolicyError(ValueError):
    def __init__(self, code: str, message: str, **detail: Any):
        super().__init__(message)
        self.code, self.detail = code, detail

    def as_issue(self) -> dict:
        return {'severity': 'error', 'code': self.code, 'message': str(self), **self.detail}


def digest(data: bytes) -> str:
    return 'sha256:' + hashlib.sha256(data).hexdigest()


def file_digest(path: Path) -> str:
    return digest(Path(path).read_bytes())


def parse(data: bytes):
    return E.fromstring(data, E.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False))


def dump(root) -> bytes:
    return E.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)


def node(tag: str, **attrs):
    return E.Element(Q(tag), {Q(k): str(v) for k, v in attrs.items()})


def value(element, name='val', default=None):
    return default if element is None else element.get(Q(name), default)


def set_child(parent, tag, **attrs):
    result = parent.find(Q(tag))
    if result is None:
        result = node(tag)
        parent.append(result)
    for key, val in attrs.items():
        result.set(Q(key), str(val))
    return result


def props(p):
    result = p.find(Q('pPr'))
    if result is None:
        result = node('pPr'); p.insert(0, result)
    return result


def visible(element) -> str:
    result = []
    for n in element.iter():
        if any(a.tag in {Q('del'), Q('moveFrom')} for a in n.iterancestors()):
            continue
        if n.tag == Q('t'):
            result.append(n.text or '')
        elif n.tag == Q('tab'):
            result.append('\t')
        elif n.tag in {Q('br'), Q('cr')}:
            result.append('\n')
    return ''.join(result)


def run(text='', rpr=None):
    r = node('r')
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    if text:
        t = node('t'); t.text = str(text); t.set(XMLSPACE, 'preserve'); r.append(t)
    return r


def seq_field(instruction: str, cached: str, hidden=False):
    f = node('fldSimple', instr=' ' + instruction.strip() + ' ', dirty='true')
    rp = node('rPr') if hidden else None
    if hidden:
        rp.append(node('vanish'))
    r = run(cached, rp)
    if hidden and not cached:
        r.append(node('t'))
    f.append(r)
    return f


def fields(p) -> list[dict]:
    """Read simple/complex fields including their cached display offsets."""
    result, stack, offset = [], [], 0
    for n in p.iter():
        if n.tag == Q('fldSimple'):
            result.append({'node': n, 'instruction': value(n, 'instr', ''),
                           'start': offset, 'end': offset + len(visible(n)),
                           'cached': visible(n), 'complete': True,
                           'locked': value(n, 'fldLock', '0').lower() not in FALSE})
        elif n.tag == Q('fldChar'):
            kind = value(n, 'fldCharType')
            if kind == 'begin':
                stack.append({'node': n, 'instruction': '', 'start': offset,
                              'cached': '', 'complete': False, 'separate': False,
                              'locked': value(n, 'fldLock', '0').lower() not in FALSE})
            elif kind == 'separate' and stack:
                stack[-1]['separate'] = True; stack[-1]['start'] = offset
            elif kind == 'end' and stack:
                f = stack.pop(); f['end'] = offset; f['complete'] = f.pop('separate'); result.append(f)
        elif n.tag == Q('instrText') and stack:
            stack[-1]['instruction'] += n.text or ''
        elif n.tag in {Q('t'), Q('tab'), Q('br'), Q('cr')}:
            txt = n.text or '' if n.tag == Q('t') else ('\t' if n.tag == Q('tab') else '\n')
            offset += len(txt)
            for f in stack:
                if f['separate']:
                    f['cached'] += txt
    for f in stack:
        f['end'] = offset; result.append(f)
    return result


def command(f) -> str:
    tokens = f['instruction'].strip().split(None, 1)
    return tokens[0].upper() if tokens else ''


def atomic_children(p):
    """Complex fields form indivisible atoms. Never leave stale field code behind."""
    children = [c for c in p if c.tag != Q('pPr')]
    groups, i = [], 0
    while i < len(children):
        c = children[i]
        starts = c.xpath('.//w:fldChar[@w:fldCharType="begin"]', namespaces=NS)
        if starts:
            group, depth = [], 0
            while i < len(children):
                current = children[i]; group.append(current); i += 1
                for mark in current.iter(Q('fldChar')):
                    if value(mark, 'fldCharType') == 'begin': depth += 1
                    elif value(mark, 'fldCharType') == 'end': depth -= 1
                if depth == 0: break
            if depth:
                raise PolicyError('field-unclosed', '域结构未闭合，不能按普通文本替换。')
            groups.append((group, 'field'))
        else:
            groups.append(([c], 'field' if c.tag == Q('fldSimple') else 'plain'))
            i += 1
    offset = 0
    for group, kind in groups:
        text = ''.join(visible(c) for c in group)
        yield group, kind, offset, offset + len(text), text
        offset += len(text)


def slice_plain_run(r, start, end):
    """Copy a text-only run fragment while retaining its original run properties."""
    if r.tag != Q('r') or any(c.tag not in {Q('rPr'), Q('t'), Q('tab'), Q('br'), Q('cr')} for c in r):
        raise PolicyError('unsafe-text-span', '待替换文字与超链接、对象或内容控件混排，需精确定位。')
    result = E.Element(r.tag, dict(r.attrib))
    rp = r.find(Q('rPr'))
    if rp is not None: result.append(copy.deepcopy(rp))
    pos = 0
    for child in r:
        if child.tag == Q('rPr'): continue
        txt = child.text or '' if child.tag == Q('t') else ('\t' if child.tag == Q('tab') else '\n')
        lo, hi = max(start - pos, 0), min(end - pos, len(txt)); pos += len(txt)
        if lo >= hi: continue
        n = copy.deepcopy(child)
        if n.tag == Q('t'): n.text = txt[lo:hi]; n.set(XMLSPACE, 'preserve')
        result.append(n)
    return result


def replace_span(p, start: int, end: int, replacements: list):
    """Replace a bounded display range, removing whole fields when selected.

    Preserves title runs and bookmark boundaries. Partial field replacements and
    ranges crossing non-text objects are rejected instead of flattening them.
    """
    if not 0 <= start <= end <= len(visible(p)):
        raise PolicyError('invalid-span', '文字替换范围无效。')
    atoms = list(atomic_children(p))
    new, inserted = [], False
    kept_starts, kept_ends = [], []
    for group, kind, a, b, txt in atoms:
        if b <= start and not (a == b == start):
            new.extend(group); continue
        if a >= end and not (start == end == a and not inserted):
            if not inserted:
                new.extend(kept_starts + replacements + kept_ends); inserted = True
            new.extend(group); continue
        if a == b:
            for n in group:
                if n.tag == Q('bookmarkStart'): kept_starts.append(n)
                elif n.tag == Q('bookmarkEnd'): kept_ends.append(n)
                else: new.append(n)
            continue
        lo, hi = max(start, a) - a, min(end, b) - a
        if kind == 'field':
            if not (start <= a and end >= b):
                if start == end == a:
                    if not inserted:
                        new.extend(kept_starts + replacements + kept_ends); inserted = True
                    new.extend(group); continue
                raise PolicyError('partial-field-span', '不能只修改域结果的一部分；必须替换完整引用/编号域。')
            for c in group:
                if c.tag == Q('bookmarkStart'): kept_starts.append(c)
                if c.tag == Q('bookmarkEnd'): kept_ends.append(c)
        else:
            c = group[0]
            if lo > 0: new.append(slice_plain_run(c, 0, lo))
        if not inserted:
            new.extend(kept_starts + replacements); kept_starts = []; inserted = True
        if hi < len(txt):
            if kind == 'field': raise PolicyError('partial-field-span', '编号域边界不完整。')
            new.append(slice_plain_run(group[0], hi, len(txt)))
    if not inserted: new.extend(kept_starts + replacements)
    new.extend(kept_ends)
    for c in list(p):
        if c.tag != Q('pPr'): p.remove(c)
    for c in new: p.append(c)


def manual_prefix(text: str, level: int) -> dict | None:
    # Dates and quantities can be the substantive beginning of a heading.
    if re.match(r'^\s*(?:(?:19|20)\d{2}\s*(?:年|年度)|(?:19|20)\d{2}(?:[.．/\-]\d{1,2}){1,2}(?:\s|$)|\d+\s+(?:万元|亿元|万亩|平方米|公里|平方公里)(?:\s|[^0-9]))', text):
        return None
    # Explicit chapter wording is a source defect when it disagrees with template.
    m = re.match(r'^\s*第\s*([0-9' + CN + r']+)\s*(章|节|篇|部分)([ \t\u3000]*)', text)
    if m:
        return {'end': m.end(), 'fmt': 'decimal' if m[1].isdecimal() else 'chineseCounting',
                'pattern': m[0][:m.start(1)] + '%' + str(level + 1) + m[0][m.end(1):m.end(2)], 'suff': suffix(m[3])}
    m = re.match(r'^\s*([（(]?)([' + CN + r']+)([）).．、])([ \t\u3000]*)', text)
    if m:
        return {'end': m.end(), 'fmt': 'chineseCounting', 'pattern': m[1] + '%' + str(level + 1) + m[3], 'suff': suffix(m[4])}
    m = re.match(r'^\s*([（(]?)(\d+(?:[.．]\d+)*)([）).．、]?)([ \t\u3000]*)', text)
    if m and (m[3] or m[4] or '.' in m[2] or '．' in m[2]):
        values = re.split('[.．]', m[2])
        if len(values) not in {1, level + 1}:
            raise PolicyError('heading-level-ambiguous', '手工编号层级与标题层级不一致，不能猜改。', text=text)
        parts = ['%' + str(i + 1) for i in range(len(values))] if len(values) > 1 else ['%' + str(level + 1)]
        return {'end': m.end(), 'fmt': 'decimal', 'pattern': m[1] + ('．' if '．' in m[2] else '.').join(parts) + m[3], 'suff': suffix(m[4])}
    return None


def suffix(text):
    return 'tab' if '\t' in text else 'space' if text else 'nothing'


class Doc:
    def __init__(self, path: Path):
        self.path = Path(path)
        with zipfile.ZipFile(path) as z:
            if len(z.namelist()) != len(set(z.namelist())):
                raise PolicyError('duplicate-part', 'DOCX 存在重复部件。')
            self.files = {n: z.read(n) for n in z.namelist()}
        for part in ('word/document.xml', 'word/styles.xml', '[Content_Types].xml'):
            if part not in self.files: raise PolicyError('missing-part', 'DOCX 缺少必要部件。', part=part)
        self.roots = {}
        self.document = self.root('word/document.xml')
        self.styles = self.root('word/styles.xml')
        self.numbering = self.root('word/numbering.xml', 'numbering')
        self.changed = set()
        self.refresh()

    def root(self, part, create=None):
        if part not in self.roots:
            if part in self.files: self.roots[part] = parse(self.files[part])
            elif create: self.roots[part] = E.Element(Q(create), nsmap={'w': W})
            else: raise PolicyError('missing-part', 'DOCX 部件不存在。', part=part)
        return self.roots[part]

    def refresh(self):
        self.style_map = {value(s, 'styleId'): s for s in self.styles.findall(Q('style'))}
        self.paragraphs = self.document.xpath('.//w:p[not(ancestor::w:del or ancestor::w:moveFrom)]', namespaces=NS)
        self.pindex = {p: i for i, p in enumerate(self.paragraphs, 1)}
        self.order = {n: i for i, n in enumerate(self.document.iter())}

    def style_id(self, p):
        return value(p.find('w:pPr/w:pStyle', NS))

    def chain(self, sid):
        result, seen = [], set()
        while sid:
            if sid in seen: raise PolicyError('style-cycle', '样式继承循环。', style_id=sid)
            if sid not in self.style_map: break
            seen.add(sid); s = self.style_map[sid]; result.append(s); sid = value(s.find(Q('basedOn')))
        return list(reversed(result))

    def style_name(self, p):
        s = self.style_map.get(self.style_id(p))
        return value(s.find(Q('name')), default='') if s is not None else ''

    def effective(self, sid, kind):
        result = node(kind)
        defaults = self.styles.find('w:docDefaults/w:' + kind + 'Default/w:' + kind, NS)
        sources = ([defaults] if defaults is not None else []) + [s.find(Q(kind)) for s in self.chain(sid)]
        for source in sources:
            if source is None: continue
            for c in source:
                old = result.find(c.tag)
                if old is None: result.append(copy.deepcopy(c)); continue
                # Attribute-valued properties (fonts, indents, spacing) inherit per slot.
                if not len(c) and not len(old):
                    if c.tag == Q('rFonts'):
                        # A child style's explicit face overrides an inherited
                        # theme face for that slot; do not resurrect the theme
                        # when materializing template inheritance.
                        for slot in ('ascii','hAnsi','eastAsia','cs'):
                            theme_keys = [Q(slot+'Theme')]
                            if slot == 'cs':theme_keys.append(Q('cstheme'))
                            if Q(slot) in c.attrib and not any(k in c.attrib for k in theme_keys):
                                for k in theme_keys:old.attrib.pop(k,None)
                            elif any(k in c.attrib for k in theme_keys):
                                old.attrib.pop(Q(slot),None)
                    old.attrib.update(c.attrib)
                else:
                    result.replace(old, copy.deepcopy(c))
        return result

    def level(self, p):
        if p.xpath('ancestor::w:tc|ancestor::w:txbxContent', namespaces=NS): return None
        names = [value(s.find(Q('name')), default='') for s in self.chain(self.style_id(p))]
        if any(re.match(r'^(?:TOC|目录|Title$|Subtitle$|封面)', name, re.I) for name in names): return None
        direct = p.find('w:pPr/w:outlineLvl', NS)
        n = direct if direct is not None else self.effective(self.style_id(p), 'pPr').find(Q('outlineLvl'))
        if n is not None:
            lv = int(value(n)); return lv if 0 <= lv <= 8 else None
        for name in reversed(names):
            m = re.fullmatch(r'(?:Heading|标题)\s*([1-9])', name, re.I)
            if m: return int(m[1]) - 1
        return None

    def num(self, p):
        inherited = self.effective(self.style_id(p), 'pPr').find(Q('numPr'))
        direct = p.find('w:pPr/w:numPr', NS)
        nid, lv = None, None
        for n in (inherited, direct):
            if n is None: continue
            if n.find(Q('numId')) is not None: nid = value(n.find(Q('numId')))
            if n.find(Q('ilvl')) is not None: lv = int(value(n.find(Q('ilvl'))))
        if nid in {None, '0'}: return None
        instance = self.numbering.find('w:num[@w:numId="' + nid + '"]', NS)
        if instance is None: return None
        aid = value(instance.find(Q('abstractNumId')))
        abstract = self.numbering.find('w:abstractNum[@w:abstractNumId="' + str(aid) + '"]', NS)
        seen = set()
        while abstract is not None and abstract.find(Q('numStyleLink')) is not None:
            if aid in seen: raise PolicyError('numbering-cycle', '编号样式链接循环。')
            seen.add(aid)
            linked = value(abstract.find(Q('numStyleLink')))
            n = self.effective(linked, 'pPr').find('w:numPr/w:numId', NS)
            ni = self.numbering.find('w:num[@w:numId="' + str(value(n)) + '"]', NS)
            aid = value(ni.find(Q('abstractNumId'))) if ni is not None else None
            abstract = self.numbering.find('w:abstractNum[@w:abstractNumId="' + str(aid) + '"]', NS)
        if abstract is None: return None
        if lv is None:
            ids = {value(s, 'styleId') for s in self.chain(self.style_id(p))}
            matches = [int(value(n, 'ilvl')) for n in abstract.findall(Q('lvl')) if value(n.find(Q('pStyle'))) in ids]
            lv = matches[-1] if matches else (self.level(p) or 0)
        override = instance.find('w:lvlOverride[@w:ilvl="' + str(lv) + '"]', NS)
        definition = override.find(Q('lvl')) if override is not None else None
        if definition is None: definition = abstract.find('w:lvl[@w:ilvl="' + str(lv) + '"]', NS)
        if definition is None: return None
        definition = copy.deepcopy(definition)
        if override is not None and override.find(Q('startOverride')) is not None:
            set_child(definition, 'start', val=value(override.find(Q('startOverride'))))
        return {'id': nid, 'level': lv, 'definition': definition, 'abstract': abstract}

    def register(self, kind):
        part = 'word/' + kind + '.xml'
        relname = 'word/_rels/document.xml.rels'
        if relname not in self.roots:
            self.roots[relname] = parse(self.files[relname]) if relname in self.files else E.Element('{' + REL + '}Relationships', nsmap={None: REL})
        rels = self.roots[relname]
        existing = [n for n in rels if n.get('Type') == R + '/' + kind]
        if existing:
            if len(existing) != 1 or existing[0].get('Target') not in {kind + '.xml', '/' + part} or existing[0].get('TargetMode') == 'External':
                raise PolicyError('nonstandard-relation', '现有部件关系不能安全覆盖。', part=part)
        else:
            ids = {n.get('Id') for n in rels}; i = 1
            while 'rId' + str(i) in ids: i += 1
            E.SubElement(rels, '{' + REL + '}Relationship', Id='rId' + str(i), Type=R + '/' + kind, Target=kind + '.xml')
            self.changed.add(relname)
        ct = self.root('[Content_Types].xml')
        if not any(n.get('PartName') == '/' + part for n in ct):
            E.SubElement(ct, '{' + CT + '}Override', PartName='/' + part,
                         ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.' + kind + '+xml')
            self.changed.add('[Content_Types].xml')

    def save(self, out):
        out = Path(out)
        if out.resolve() == self.path.resolve(): raise PolicyError('in-place-write', '输出不能覆盖输入。')
        data = dict(self.files)
        for part in self.changed: data[part] = dump(self.roots[part])
        out.parent.mkdir(parents=True, exist_ok=True)
        fd, tmp = tempfile.mkstemp(suffix='.docx', dir=out.parent); os.close(fd)
        try:
            with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as z:
                for name, content in data.items(): z.writestr(name, content)
            with zipfile.ZipFile(tmp) as z:
                if z.testzip(): raise PolicyError('bad-output-zip', '输出 DOCX 校验失败。')
            os.replace(tmp, out)
        finally:
            if os.path.exists(tmp): os.unlink(tmp)


def role_ids(style_json: Path | None) -> dict:
    if style_json is None: return {}
    raw = json.loads(Path(style_json).read_text(encoding='utf-8-sig'))
    roles = {}
    for name, info in (raw.get('semantic_styles') or raw.get('semantic_roles') or {}).items():
        if isinstance(info, dict):
            sid = info.get('source_style_id') or info.get('style_id')
            if sid: roles[name] = str(sid)
    return roles


def fake_p(sid):
    p = node('p'); pp = node('pPr'); pp.append(node('pStyle', val=sid)); p.append(pp); return p


def level_signature(lvl):
    out = copy.deepcopy(lvl)
    out.attrib.pop(Q('tplc'), None)
    for n in list(out):
        if n.tag in {Q('pStyle'), Q('rPr'), Q('pPr')}:
            out.remove(n)
    # Include suffix defaults: a missing suffix in OOXML means tab, not space.
    if out.find(Q('suff')) is None: set_child(out, 'suff', val='tab')
    if out.find(Q('start')) is None: set_child(out, 'start', val='1')
    return sorted((E.QName(n).localname, tuple(sorted(n.attrib.items()))) for n in out)


def template_contract(template: Path, levels: set[int], style_json: Path | None = None):
    src = Doc(template); roles = role_ids(style_json); definitions, styles = {}, {}
    examples = defaultdict(list)
    for p in src.paragraphs:
        lv = src.level(p)
        if lv is not None and visible(p).strip(): examples[lv].append(p)
    for lv in sorted(levels):
        sid = roles.get('heading' + str(lv + 1))
        if sid and sid not in src.style_map:
            raise PolicyError('template-role-invalid', '模板声明的标题样式不存在。', style_id=sid)
        if not sid:
            used = {src.style_id(p) for p in examples[lv] if src.style_id(p)}
            candidates = used or {k for k, s in src.style_map.items() if value(s, 'type') == 'paragraph' and src.level(fake_p(k)) == lv}
            signatures = {(E.tostring(src.effective(k, 'pPr')), E.tostring(src.effective(k, 'rPr'))) for k in candidates}
            if len(signatures) > 1:
                raise PolicyError('template-heading-conflict', '模板同层级有多套不同样式，需明确选择，不能按频次猜测。', level=lv + 1)
            sid = sorted(candidates)[0] if candidates else None
        if sid is None:
            raise PolicyError('template-heading-missing', '模板缺少当前层级的标题约定。', level=lv + 1)
        samples = [p for p in examples[lv] if src.style_id(p) == sid]
        numbered = [src.num(p) for p in samples + [fake_p(sid)]]
        numbered = [n for n in numbered if n is not None]
        if not numbered:
            # Some style-only templates link a level via lvl/pStyle without a
            # sample paragraph or a numPr on the style. Read that native link.
            for abstract in src.numbering.findall(Q('abstractNum')):
                for definition in abstract.findall(Q('lvl')):
                    if value(definition.find(Q('pStyle'))) == sid and int(value(definition,'ilvl')) == lv:
                        numbered.append({'definition':copy.deepcopy(definition)})
        if numbered:
            signatures = {repr(level_signature(n['definition'])) for n in numbered}
            if len(signatures) != 1:
                raise PolicyError('template-numbering-conflict', '模板同层级编号定义冲突。', level=lv + 1)
            definition = copy.deepcopy(numbered[0]['definition'])
        else:
            prefixes = [manual_prefix(visible(p), lv) for p in samples]
            prefixes = [p for p in prefixes if p]
            shapes = {(p['fmt'], p['pattern'], p['suff']) for p in prefixes}
            if len(shapes) != 1:
                raise PolicyError('template-numbering-missing', '模板没有可唯一确定的编号定义或示例，不能从待修文档反推。', level=lv + 1)
            fmt, pattern, suff = next(iter(shapes))
            definition = node('lvl', ilvl=lv)
            for tag, v in [('start', '1'), ('numFmt', fmt), ('suff', suff), ('lvlText', pattern), ('lvlJc', 'left')]:
                definition.append(node(tag, val=v))
        fmt = value(definition.find(Q('numFmt')))
        pattern = value(definition.find(Q('lvlText')), default='')
        if fmt in {None, 'none', 'bullet'} or not re.search(r'%[1-9]', pattern):
            raise PolicyError('template-numbering-invalid', '模板标题编号不是有效的数字序列。', level=lv + 1)
        definitions[lv], styles[lv] = definition, sid
    caption_sid = roles.get('caption')
    if not caption_sid:
        used_caps = {src.style_id(p) for p in src.paragraphs if re.search(r'caption|题注|图题|表题', src.style_name(p), re.I)}
        if len(used_caps) == 1: caption_sid = next(iter(used_caps))
        else:
            caps = [k for k, s in src.style_map.items() if value(s, 'type') == 'paragraph' and re.search(r'caption|题注|图题|表题', value(s.find(Q('name')), default=''), re.I)]
            if len(caps) == 1: caption_sid = caps[0]
            elif '30' in caps and Path(template).name in {'default-template.docx', 'default-technical-bid.docx'}:
                # Existing approved bundle contract; never applies to an uploaded template.
                caption_sid = '30'
    return {'source': src, 'styles': styles, 'levels': definitions, 'caption_style': caption_sid,
            'template_sha256': file_digest(template)}


def ordered_properties(parent):
    """Keep generated property children in the OOXML schema order."""
    orders = {
        'style': 'name aliases basedOn next link autoRedefine hidden uiPriority semiHidden unhideWhenUsed qFormat locked personal personalCompose personalReply rsid pPr rPr tblPr trPr tcPr tblStylePr',
        'pPr': 'pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange',
        'rPr': 'rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath rPrChange',
        'lvl': 'start numFmt lvlRestart pStyle isLgl suff lvlText lvlPicBulletId legacy lvlJc pPr rPr',
    }
    order=orders.get(E.QName(parent).localname)
    if order:
        positions={name:i for i,name in enumerate(order.split())}
        children=list(parent)
        for child in children:parent.remove(child)
        for child in sorted(children,key=lambda c:positions.get(E.QName(c).localname,len(positions))):parent.append(child)
    for child in parent:
        if E.QName(child).localname in {'pPr','rPr'}:ordered_properties(child)
    return parent


def materialized_style(source: Doc, sid: str, target_id: str):
    s = copy.deepcopy(source.style_map[sid]); s.set(Q('styleId'), target_id)
    # Materialize inheritance locally: do not replace global defaults or theme colors.
    for tag in ('basedOn', 'link', 'next', 'pPr', 'rPr'):
        for child in list(s.findall(Q(tag))): s.remove(child)
    for kind in ('pPr', 'rPr'):
        effective = source.effective(sid, kind)
        for n in list(effective.findall(Q('numPr'))): effective.remove(n)
        if len(effective): s.append(effective)
    # Resolve only theme FONT slots from the template; no color palette is copied.
    theme_name = next((n for n in source.files if n.startswith('word/theme/') and n.endswith('.xml')), None)
    if theme_name:
        theme = parse(source.files[theme_name]); rfonts = s.find('w:rPr/w:rFonts', NS)
        if rfonts is not None:
            for slot in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
                themed = rfonts.get(Q(slot + 'Theme')) or (rfonts.get(Q('cstheme')) if slot == 'cs' else None)
                if not themed: continue
                prefix = 'major' if themed.startswith('major') else 'minor'
                branch = theme.find('.//a:fontScheme/a:' + prefix + 'Font', NS)
                if branch is None: continue
                if slot == 'eastAsia':
                    face = branch.find('a:font[@script="Hans"]', NS)
                    if face is None: face = branch.find('a:ea', NS)
                elif slot == 'cs': face = branch.find('a:cs', NS)
                else: face = branch.find('a:latin', NS)
                if face is not None and face.get('typeface'):
                    rfonts.set(Q(slot), face.get('typeface')); rfonts.attrib.pop(Q(slot + 'Theme'), None)
                    if slot == 'cs':rfonts.attrib.pop(Q('cstheme'),None)
    return ordered_properties(s)


def install_style(doc: Doc, contract, sid: str, role: str):
    existing = doc.style_map.get(sid)
    expected_level = int(role[7:]) - 1 if role.startswith('heading') else None
    safe = existing is None or (role.startswith('heading') and doc.level(fake_p(sid)) == expected_level) or (role == 'caption' and re.search(r'caption|题注|图题|表题', value(existing.find(Q('name')), default=''), re.I))
    target_id = sid if safe else 'DLR_' + role
    new = materialized_style(contract['source'], sid, target_id)
    if target_id != sid:
        set_child(new, 'name', val='DLR ' + value(new.find(Q('name')), default=role))
    old = doc.style_map.get(target_id)
    if old is not None: doc.styles.replace(old, new)
    else: doc.styles.append(new)
    next_id = value(contract['source'].style_map[sid].find(Q('next')))
    if next_id and next_id in contract['source'].style_map and next_id != sid:
        # Next is an editing convention from the template. A private materialized
        # style avoids mutating the document-wide Normal/body baseline.
        next_target = 'DLRNext_' + role
        expected_next = materialized_style(contract['source'], next_id, next_target)
        set_child(expected_next, 'name', val='DLR next ' + role)
        old_next = doc.style_map.get(next_target)
        if old_next is not None: doc.styles.replace(old_next, expected_next)
        else: doc.styles.append(expected_next)
        set_child(new, 'next', val=next_target)
    elif next_id == sid:
        set_child(new, 'next', val=target_id)
    ordered_properties(new)
    doc.changed.add('word/styles.xml'); doc.refresh()
    return target_id


def set_pstyle(p, sid):
    pp = props(p); ps = pp.find(Q('pStyle'))
    if ps is None: ps = node('pStyle'); pp.insert(0, ps)
    ps.set(Q('val'), sid)


def normalize_headings(doc: Doc, contract, changes):
    records = [(p, doc.level(p)) for p in doc.paragraphs if doc.level(p) is not None and visible(p).strip()]
    if not records: return
    target_styles = {lv: install_style(doc, contract, sid, 'heading' + str(lv + 1)) for lv, sid in contract['styles'].items()}
    name = 'DLRTemplate_' + contract['template_sha256'][-20:]
    abstract = next((n for n in doc.numbering.findall(Q('abstractNum')) if value(n.find(Q('name'))) == name), None)
    if abstract is None:
        aid = max((int(value(n, 'abstractNumId')) for n in doc.numbering.findall(Q('abstractNum'))), default=-1) + 1
        abstract = node('abstractNum', abstractNumId=aid)
        abstract.append(node('multiLevelType', val='multilevel')); abstract.append(node('name', val=name))
        insert_at = next((i for i, n in enumerate(doc.numbering) if n.tag == Q('num')), len(doc.numbering))
        doc.numbering.insert(insert_at, abstract)
    else:
        aid = int(value(abstract, 'abstractNumId'))
        for lv in list(abstract.findall(Q('lvl'))): abstract.remove(lv)
    for lv, definition in sorted(contract['levels'].items()):
        actual = copy.deepcopy(definition); actual.set(Q('ilvl'), str(lv))
        set_child(actual, 'pStyle', val=target_styles[lv]); abstract.append(ordered_properties(actual))
    inst = next((n for n in doc.numbering.findall(Q('num')) if value(n.find(Q('abstractNumId'))) == str(aid)), None)
    if inst is None:
        nid = max((int(value(n, 'numId')) for n in doc.numbering.findall(Q('num'))), default=0) + 1
        inst = node('num', numId=nid); inst.append(node('abstractNumId', val=aid)); doc.numbering.append(inst)
    nid = value(inst, 'numId')
    # The managed instance is continuous; never retain stale per-paragraph restarts.
    for ov in list(inst.findall(Q('lvlOverride'))): inst.remove(ov)
    for lv, sid in target_styles.items():
        style = doc.style_map[sid]
        pp = style.find(Q('pPr'))
        if pp is None: pp = node('pPr'); style.append(pp)
        for np in list(pp.findall(Q('numPr'))): pp.remove(np)
        np = node('numPr'); np.append(node('ilvl', val=lv)); np.append(node('numId', val=nid)); pp.insert(0, np)
        ordered_properties(pp)
    for p, lv in records:
        before = visible(p); prefix = manual_prefix(before, lv)
        if prefix:
            replace_span(p, 0, prefix['end'], [])
        set_pstyle(p, target_styles[lv])
        pp = props(p)
        for tag in ('numPr', 'spacing', 'ind', 'jc'):
            for n in list(pp.findall(Q(tag))): pp.remove(n)
        np = node('numPr'); np.append(node('ilvl', val=lv)); np.append(node('numId', val=nid))
        # Keep numPr before outlineLvl/sectPr and after keep flags.
        before_tags = {Q(x) for x in ('pBdr', 'shd', 'tabs', 'spacing', 'ind', 'jc', 'outlineLvl', 'rPr', 'sectPr')}
        position = next((i for i, child in enumerate(pp) if child.tag in before_tags), len(pp))
        pp.insert(position, np)
        set_child(pp, 'outlineLvl', val=lv)
        for rp in p.findall('.//w:rPr', NS):
            for tag in ('rFonts', 'sz', 'szCs', 'b', 'bCs', 'i', 'iCs'):
                for n in list(rp.findall(Q(tag))): rp.remove(n)
        changes.append({'kind': 'heading-template', 'paragraph': doc.pindex.get(p), 'before': before,
                        'format': value(contract['levels'][lv].find(Q('numFmt'))),
                        'pattern': value(contract['levels'][lv].find(Q('lvlText'))),
                        'suffix': value(contract['levels'][lv].find(Q('suff')), default='tab')})
    doc.changed.update({'word/document.xml', 'word/numbering.xml'}); doc.register('numbering'); doc.refresh()


def relationship_map(doc: Doc):
    part = 'word/_rels/document.xml.rels'
    root = doc.root(part) if part in doc.files or part in doc.roots else []
    return {n.get('Id'): n for n in root}


def caption_candidate(doc: Doc, p):
    if p is None or p.tag != Q('p') or p.xpath('.//w:drawing|.//w:pict', namespaces=NS): return False
    return bool(CAP.match(visible(p)) or re.search(r'caption|题注|图题|表题', doc.style_name(p), re.I) or
                (re.match(r'^\s*(图|表|Figure|Table)', visible(p), re.I) and any(command(f) == 'SEQ' for f in fields(p))))


def adjacent(element, forward=True):
    current = element.getnext() if forward else element.getprevious()
    blanks = 0
    while current is not None and current.tag == Q('p') and not visible(current).strip() and not current.xpath('.//w:drawing|.//w:pict', namespaces=NS):
        blanks += 1
        if blanks > 2: return None
        current = current.getnext() if forward else current.getprevious()
    return current


def inventory(doc: Doc) -> list[dict]:
    objects, seen, rels = [], set(), relationship_map(doc)
    counters = defaultdict(int)
    for el in doc.document.iter():
        kind, owner = None, el
        if el.tag == Q('tbl'):
            kind = 'table'
        elif el.tag in {Q('drawing'), Q('pict')} and not any(a.tag in {Q('drawing'), Q('pict')} for a in el.iterancestors()):
            owner = next((a for a in el.iterancestors() if a.tag == Q('p') and not a.xpath('ancestor::w:txbxContent', namespaces=NS)), None)
            if owner is None or owner in seen: continue
            kind = 'figure'; seen.add(owner)
        if not kind: continue
        counters[kind] += 1
        ident = ('F' if kind == 'figure' else 'T') + str(counters[kind]).zfill(4)
        paths, media, external = [], {}, False
        if kind == 'figure':
            for n in owner.iter():
                for key in ('{' + R + '}embed', '{' + R + '}link', '{' + R + '}id'):
                    rid = n.get(key)
                    if not rid: continue
                    relation = rels.get(rid)
                    if relation is None: continue
                    if relation.get('TargetMode') == 'External': external = True; continue
                    target = relation.get('Target', '')
                    path = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('word', target))
                    if path in doc.files and '/media/' in path:
                        paths.append(path); media[path] = digest(doc.files[path])
            native_shapes = [n for n in owner.iter() if E.QName(n).localname in {'wsp', 'sp'} or (n.tag == '{' + V + '}shape' and n.find('{' + V + '}imagedata') is None)]
            if media and not native_shapes:
                object_hash = digest(json.dumps(sorted(media.values())).encode())
            elif media:
                object_hash = digest(json.dumps(sorted(media.values())).encode() + b''.join(E.tostring(n, method='c14n') for n in native_shapes))
            else:
                content = b''.join(E.tostring(n, method='c14n') for n in owner if n.tag in {Q('r'), Q('pict'), Q('drawing')})
                object_hash = digest(content)
            alt = [n.get(k) for n in owner.iter() for k in ('descr', 'title') if n.get(k)]
        else:
            rows = [[visible(c).strip() for c in row.findall(Q('tc'))] for row in el.findall(Q('tr'))]
            object_hash = digest(json.dumps(rows, ensure_ascii=False).encode())
            alt = [value(n) for n in el.findall('w:tblPr/w:tblCaption', NS) if value(n)]
        preferred = adjacent(owner, kind == 'figure'); opposite = adjacent(owner, kind != 'figure')
        caption = preferred if caption_candidate(doc, preferred) else opposite if caption_candidate(doc, opposite) else None
        if caption is not None:
            label = re.match(r'^\s*(图|表|Figure|Table)', visible(caption), re.I)
            if label and ((label[1].lower() in {'图', 'figure'}) != (kind == 'figure')): caption = None
        p = owner if owner.tag == Q('p') else next(iter(owner.findall('.//w:p', NS)), None)
        context = []
        before = owner.getprevious()
        for _ in range(3):
            if before is None: break
            if before.tag == Q('p') and visible(before).strip(): context.append(visible(before)[:250])
            before = before.getprevious()
        objects.append({'id': ident, 'kind': kind, 'element': owner, 'hash': object_hash,
                        'position': doc.order[owner], 'paragraph': doc.pindex.get(p), 'caption': caption,
                        'media': media, 'paths': sorted(set(paths)), 'external': external,
                        'alt': alt, 'context': list(reversed(context)),
                        'summary': visible(owner)[:300] if kind == 'table' else ''})
    return objects


def public_inventory(doc: Doc):
    return [{k: v for k, v in obj.items() if k not in {'element', 'caption'}} |
            {'caption_text': visible(obj['caption']) if obj['caption'] is not None else None,
             'caption_missing': obj['caption'] is None} for obj in inventory(doc)]


def load_object_plan(path, doc: Doc, for_repair=True):
    if path is None: return {}
    data = path if isinstance(path, dict) else json.loads(Path(path).read_text(encoding='utf-8-sig'))
    if not isinstance(data, dict) or set(data) - {'version', 'source_sha256', 'objects'}:
        raise PolicyError('object-plan-invalid', '图表语义计划含未知字段。')
    if for_repair and data.get('source_sha256') != file_digest(doc.path):
        raise PolicyError('object-plan-stale', '图表语义计划不属于当前输入文件，必须重新定位。')
    result = {}
    actual = {o['id']: o for o in inventory(doc)}
    for row in data.get('objects', []):
        if not isinstance(row, dict) or set(row) - {'id', 'object_sha256', 'title', 'exempt', 'reason', 'caption_paragraph', 'caption_text'}:
            raise PolicyError('object-plan-invalid', '图表语义项含未知字段。')
        ident = row.get('id')
        if ident in result or ident not in actual or row.get('object_sha256') != actual[ident]['hash']:
            raise PolicyError('object-plan-stale', '图表对象不存在、重复或内容已变化。', object_id=ident)
        if row.get('exempt') and not str(row.get('reason', '')).strip():
            raise PolicyError('exemption-without-reason', '装饰图/布局表豁免必须说明对象用途。', object_id=ident)
        result[ident] = row
    return result


def chapter_map(doc: Doc):
    result, chapter = {}, 0
    for n in doc.document.iter():
        if n.tag == Q('p') and doc.level(n) == 0 and visible(n).strip(): chapter += 1
        result[n] = chapter
    return result


def caption_info(p):
    text = visible(p); fs = fields(p); seqs = [f for f in fs if command(f) == 'SEQ']
    if seqs:
        if len(seqs) != 1 or not seqs[0]['complete']:
            raise PolicyError('caption-field-invalid', '题注的 SEQ 域不唯一或未闭合。')
        label = re.match(r'^\s*(图|表|Figure|Table)\s*', text, re.I)
        if not label:
            raise PolicyError('caption-label-ambiguous', '题注域存在，但图表标签不明确。')
        end = seqs[0]['end']
        while end < len(text) and text[end].isspace(): end += 1
        m = CAP.match(text)
        return {'title': text[end:].strip(), 'end': end, 'old_number': m[2] if m else None, 'fields': fs}
    m = CAP.match(text)
    if m: return {'title': text[m.end():].strip(), 'end': m.end(), 'old_number': m[2], 'fields': fs}
    if ERROR_TEXT.search(text):
        raise PolicyError('caption-title-ambiguous', '损坏的静态题注需根据图表内容确认题名，不能猜切错误文本。')
    return {'title': text.strip(), 'end': 0, 'old_number': None, 'fields': fs}


def choose_title(obj, row):
    if row and str(row.get('title', '')).strip(): title = row['title'].strip()
    else:
        titles = [t.strip() for t in obj['alt'] if t.strip() and not re.search(r'\.(png|jpe?g|svg|emf|wmf)$', t, re.I)]
        title = titles[0] if len(set(titles)) == 1 else None
    if not title:
        raise PolicyError('caption-title-required', '缺少题注：Agent 必须查看对象与上下文后填写题名。', object_id=obj['id'], context=obj['context'], summary=obj['summary'])
    title = CAP.sub('', title, count=1).strip()
    if not title or ERROR_TEXT.search(title):
        raise PolicyError('caption-title-invalid', '题名不能为空或包含域错误。', object_id=obj['id'])
    return title


def normalize_captions(doc: Doc, contract, plan: dict, changes):
    objects = inventory(doc); chapters = chapter_map(doc); old_refs = []
    if not objects: return old_refs
    sid = contract.get('caption_style')
    if sid not in contract['source'].style_map:
        raise PolicyError('template-caption-style-missing', '模板题注样式未唯一确定；请指定已确认的题注样式。')
    target_sid = install_style(doc, contract, sid, 'caption')
    counts = defaultdict(int); assigned = set()
    for obj in objects:
        row = plan.get(obj['id'], {})
        if row.get('exempt'): continue
        if obj['external']:
            raise PolicyError('linked-image-unresolved', '外链图片须先读取实际图像并确认归属。', object_id=obj['id'])
        cap = obj['caption']; owner = obj['element']
        if row.get('caption_paragraph') is not None:
            index = row['caption_paragraph']
            if type(index) is not int or not 1 <= index <= len(doc.paragraphs):
                raise PolicyError('caption-plan-location', '题注定位无效。', object_id=obj['id'])
            cap = doc.paragraphs[index - 1]
            if visible(cap) != row.get('caption_text') or cap.xpath('.//w:drawing|.//w:pict', namespaces=NS):
                raise PolicyError('caption-plan-location', '题注文字锚点与当前内容不一致。', object_id=obj['id'])
        if cap is not None and cap in assigned:
            raise PolicyError('caption-shared-ambiguous', '多个对象共用题注，须先确认是组合图还是独立图。', object_id=obj['id'])
        info = caption_info(cap) if cap is not None else None
        title = info['title'] if info and info['title'] else choose_title(obj, row)
        if row.get('title'): title = choose_title(obj, row)
        chapter = chapters.get(owner, 0)
        counts[(chapter, obj['kind'])] += 1
        seq = SEQUENCES[obj['kind']]
        cached = str(counts[(chapter, obj['kind'])])
        bookmarks = [value(b, 'name') for b in cap.findall(Q('bookmarkStart'))] if cap is not None else []
        if info and info['old_number']:
            old_refs.append({'kind': obj['kind'], 'number': info['old_number'], 'chapter': chapter,
                             'owner': owner, 'title': title, 'bookmarks': bookmarks, 'id': obj['id']})
        elif bookmarks:
            old_refs.append({'kind': obj['kind'], 'number': None, 'chapter': chapter,
                             'owner': owner, 'title': title, 'bookmarks': bookmarks, 'id': obj['id']})
        instruction = 'SEQ ' + seq + ' \\* ARABIC \\s 1'
        prefix = [run(LABELS[obj['kind']]), seq_field(instruction, cached), run(' ')]
        if cap is None:
            cap = node('p'); cap.extend(prefix + [run(title)])
            if obj['kind'] == 'figure': owner.addnext(cap)
            else: owner.addprevious(cap)
            action = '补充独立 Word 题注'
        else:
            assigned.add(cap)
            if not info['title'] or (row.get('title') and info['title'] != title):
                # Explicit semantic plan authorizes only this caption's title.
                replace_span(cap, 0, len(visible(cap)), prefix + [run(title)])
            elif info['end'] == 0:
                replace_span(cap, 0, 0, prefix)
            else:
                replace_span(cap, 0, info['end'], prefix)
            if cap.getparent() is not owner.getparent():
                raise PolicyError('caption-container-ambiguous', '题注与对象不在同一容器，需确认移动位置。', object_id=obj['id'])
            if obj['kind'] == 'figure': owner.addnext(cap)
            else: owner.addprevious(cap)
            action = '题注改为分章独立序列，移除章节前缀/失效 STYLEREF'
        set_pstyle(cap, target_sid)
        pp = props(cap)
        for tag in ('numPr', 'ind', 'spacing', 'jc'):
            for n in list(pp.findall(Q(tag))): pp.remove(n)
        set_child(pp, 'keepNext', val='0' if obj['kind'] == 'figure' else '1')
        set_child(pp, 'keepLines', val='1')
        for rp in cap.findall('.//w:rPr', NS):
            for tag in ('rFonts', 'sz', 'szCs', 'b', 'bCs', 'i', 'iCs'):
                for n in list(rp.findall(Q(tag))): rp.remove(n)
        if obj['kind'] == 'figure': set_child(props(owner), 'keepNext', val='1')
        changes.append({'kind': 'caption', 'object_id': obj['id'], 'action': action,
                        'chapter': chapter, 'title': title, 'number': LABELS[obj['kind']] + cached})
    # Anchor resets to the CHAPTER, never to its first caption. Deleting/inserting
    # the first caption cannot delete or duplicate the chapter's reset boundary.
    for p in doc.paragraphs:
        if doc.level(p) != 0 or not visible(p).strip(): continue
        for f in fields(p):
            if command(f) == 'SEQ' and re.search(r'\b(?:Figure|Table|图|表)\b', f['instruction'], re.I) and re.search(r'\\r\s+0\b', f['instruction'], re.I) and re.search(r'\\h\b', f['instruction'], re.I):
                if f['node'].tag == Q('fldSimple'):
                    f['node'].getparent().remove(f['node'])
                else:
                    # Word may serialize a simple field as begin/instruction/
                    # separate/result/end runs after F9. Remove the whole reset,
                    # not merely its displayed zero; preserve the heading text.
                    matching = [(group, txt) for group, kind, _, _, txt in atomic_children(p)
                                if kind == 'field' and any(f['node'] is n for c in group for n in c.iter())]
                    if len(matching) != 1:
                        raise PolicyError('chapter-reset-complex', '隐藏重置域的完整边界不能唯一确认。')
                    group, txt = matching[0]
                    begins = sum(1 for c in group for mark in c.iter(Q('fldChar'))
                                 if value(mark, 'fldCharType') == 'begin')
                    if begins != 1 or txt != f['cached'] or not re.fullmatch(r'\s*0?\s*', txt):
                        raise PolicyError('chapter-reset-complex', '隐藏重置域与标题文字混排，不能整段移除。')
                    for c in group:
                        if c.tag in {Q('bookmarkStart'), Q('bookmarkEnd')}: continue
                        if c.tag != Q('r') or any(ch.tag not in {Q('rPr'), Q('fldChar'), Q('instrText'), Q('t')} for ch in c):
                            raise PolicyError('chapter-reset-complex', '隐藏重置域含非文本对象，需精确迁移。')
                    for c in group:
                        if c.tag not in {Q('bookmarkStart'), Q('bookmarkEnd')}: p.remove(c)
        for seq in sorted(set(SEQUENCES.values())):
            p.append(seq_field('SEQ ' + seq + ' \\r 0 \\h', '', hidden=True))
    doc.changed.add('word/document.xml'); doc.refresh()
    return old_refs


def body_references(txt):
    protected = [m.span() for m in re.finditer(r'(?:https?://|www\.|[A-Za-z]:[\\/])\S+|\S+[\\/]\S+', txt)]
    for m in REF.finditer(txt):
        if any(a <= m.start() < b for a, b in protected): continue
        if re.match(r'\.(?:png|jpe?g|gif|svg|emf|wmf|pdf|docx?|xlsx?|json|xml|csv)\b', txt[m.end():], re.I): continue
        yield m


def relative_references(doc: Doc, old_refs, changes):
    if not old_refs: return
    caps = {o['caption'] for o in inventory(doc) if o['caption'] is not None}
    chapters = chapter_map(doc)
    for p in list(doc.paragraphs):
        if p in caps or doc.level(p) is not None or p.xpath('ancestor::w:txbxContent', namespaces=NS): continue
        if re.match(r'^(?:TOC|目录)', doc.style_name(p), re.I): continue
        txt = visible(p); edits = []
        for m in body_references(txt):
            kind = 'figure' if m[1].lower() in {'图', 'figure'} else 'table'
            candidates = [r for r in old_refs if r['kind'] == kind and r['number'] == m[2]]
            local = [r for r in candidates if r['chapter'] == chapters.get(p, 0)]
            candidates = local or candidates
            if len(candidates) != 1:
                raise PolicyError('reference-target-ambiguous', '正文图表引用不能唯一定位；不得用“相关图表”掩盖歧义。', paragraph=doc.pindex[p], reference=m[0])
            target = candidates[0]
            direction = '下' if doc.order[target['owner']] > doc.order[p] else '上'
            # Relative references are only safe near the target and without a
            # competing figure/table between the referring paragraph and target.
            lo, hi = sorted((doc.order[p], doc.order[target['owner']]))
            competitors = [o for o in inventory(doc) if o['kind'] == kind and lo < o['position'] < hi]
            if competitors:
                raise PolicyError('reference-not-adjacent', '引用跨越其他同类图表，需要 Agent 调整位置或用明确题名定位。', paragraph=doc.pindex[p], reference=m[0], target=target['id'])
            edits.append((m.start(), m.end(), direction + LABELS[kind]))
        # REF fields may display only a bare number (with a surrounding 图/表).
        # Number-looking text above already covers that case. A complete REF
        # displaying the label is likewise replaced as one field by replace_span.
        for f in fields(p):
            if command(f) != 'REF': continue
            tokens = re.findall(r'"([^"]+)"|([^\s]+)', f['instruction'].strip())
            parts = [a or b for a, b in tokens]
            if len(parts) < 2: continue
            matched = [r for r in old_refs if parts[1] in r['bookmarks']]
            if not matched: continue
            if any(a <= f['start'] and b >= f['end'] for a, b, _ in edits): continue
            if len(matched) != 1:
                raise PolicyError('reference-target-ambiguous', '书签对应多个图表。', paragraph=doc.pindex[p])
            target = matched[0]; kind = target['kind']
            start, end = f['start'], f['end']
            prefix = re.search(r'(图|表|Figure|Table)\s*$', txt[:start], re.I)
            if prefix: start = prefix.start()
            direction = '下' if doc.order[target['owner']] > doc.order[p] else '上'
            edits.append((start, end, direction + LABELS[kind]))
        for start, end, replacement in sorted(set(edits), reverse=True):
            replace_span(p, start, end, [run(replacement)])
            changes.append({'kind': 'relative-reference', 'paragraph': doc.pindex[p], 'before': txt[start:end], 'after': replacement})
    doc.changed.add('word/document.xml'); doc.refresh()


def required_levels(doc: Doc):
    levels = {doc.level(p) for p in doc.paragraphs if doc.level(p) is not None and visible(p).strip()}
    # Include ancestor levels: a sublevel pattern may refer to them.
    return set(range(max(levels) + 1)) if levels else set()


def audit_document(path: Path, template: Path, style_json: Path | None = None, object_plan=None):
    doc = Doc(path); contract = template_contract(template, required_levels(doc), style_json)
    plan = load_object_plan(object_plan, doc, for_repair=False) if object_plan else {}
    issues, chapters, counts = [], chapter_map(doc), defaultdict(int)
    def bad(code, message, **data): issues.append({'severity': 'error', 'code': code, 'message': message, **data})
    for p in doc.paragraphs:
        lv = doc.level(p)
        if lv is None or not visible(p).strip(): continue
        actual = doc.num(p)
        if not actual or actual['level'] != lv:
            bad('heading-not-automatic', '标题没有有效的对应层级自动编号。', paragraph=doc.pindex[p]); continue
        if level_signature(actual['definition']) != level_signature(contract['levels'][lv]):
            bad('heading-template-numbering', '标题编号字形、前后缀、分隔符或重启方式与模板不一致。', paragraph=doc.pindex[p],
                expected={'format': value(contract['levels'][lv].find(Q('numFmt'))), 'pattern': value(contract['levels'][lv].find(Q('lvlText'))), 'suffix': value(contract['levels'][lv].find(Q('suff')), default='tab')})
        if manual_prefix(visible(p), lv): bad('heading-manual-prefix', '标题仍保留手工序号，可能显示双重编号。', paragraph=doc.pindex[p])
        expected_style = materialized_style(contract['source'], contract['styles'][lv], '_expected')
        for kind in ('pPr', 'rPr'):
            expected = expected_style.find(Q(kind))
            effective = doc.effective(doc.style_id(p), kind)
            direct = p.find(Q(kind)) if kind == 'pPr' else None
            for en in ([] if expected is None else list(expected)):
                if en.tag in {Q('numPr'), Q('keepNext'), Q('keepLines'), Q('pageBreakBefore')}:
                    continue  # valid pagination changes are reviewed on final pages
                actual_prop = direct.find(en.tag) if direct is not None else None
                if actual_prop is None: actual_prop = effective.find(en.tag)
                if actual_prop is None or any(actual_prop.get(k) != v for k, v in en.attrib.items()):
                    bad('heading-template-style', '标题的实际字体、字号或段落属性偏离模板。', paragraph=doc.pindex[p], property=E.QName(en).localname)
        if value(doc.effective(doc.style_id(p), 'pPr').find('w:numPr/w:numId', NS)) != actual['id']:
            bad('heading-style-numbering-unbound', '标题样式未绑定同一多级列表，后续新增标题可能不自动续号。', paragraph=doc.pindex[p])

    objects = inventory(doc); caps = set()
    for obj in objects:
        if plan.get(obj['id'], {}).get('exempt'): continue
        cap = obj['caption']
        if cap is None:
            bad('caption-missing', '图表缺少独立 Word 题注。', object_id=obj['id']); continue
        if cap in caps: bad('caption-shared-ambiguous', '独立对象共用了题注。', object_id=obj['id'])
        caps.add(cap)
        wanted = adjacent(obj['element'], obj['kind'] == 'figure')
        if wanted is not cap: bad('caption-position', '图题应在图下、表题应在表上。', object_id=obj['id'])
        seqs = [f for f in fields(cap) if command(f) == 'SEQ']
        if len(seqs) != 1:
            bad('caption-not-automatic', '题注缺少唯一的真实 SEQ 域。', object_id=obj['id']); continue
        f = seqs[0]; instruction = f['instruction']
        if not f['complete'] or f['locked'] or re.search(r'\\[rch]\b', instruction, re.I):
            bad('caption-sequence-invalid', '可见题注不能锁定、隐藏、逐条固定重启或只引用上一个序号。', object_id=obj['id'])
        if not re.search(r'^\s*SEQ\s+' + SEQUENCES[obj['kind']] + r'\s', instruction, re.I):
            bad('caption-sequence-mixed', '图和表必须使用各自独立的序列。', object_id=obj['id'])
        if not re.search(r'\\s\s+1\b', instruction, re.I):
            bad('caption-chapter-reset-missing', '题注序号未绑定一级章节重启。', object_id=obj['id'])
        if any(command(x) == 'STYLEREF' for x in fields(cap)):
            bad('caption-styleref-forbidden', '本次题注不允许显示章节前缀或拼接 STYLEREF。', object_id=obj['id'])
        counts[(chapters.get(obj['element'], 0), obj['kind'])] += 1
        number = counts[(chapters.get(obj['element'], 0), obj['kind'])]
        prefix = LABELS[obj['kind']] + str(number)
        if not re.match(r'^' + re.escape(prefix) + r'\s+\S', visible(cap)):
            bad('caption-cache-sequence', '题注显示应为图1/表1等章内顺序，不能保留图2.1或旧号。', object_id=obj['id'], expected=prefix)
        if f['cached'].strip() != str(number): bad('caption-cache-sequence', '题注域缓存与当前章内顺序不一致。', object_id=obj['id'])
        if ERROR_TEXT.search(visible(cap)): bad('caption-field-error', '题注仍有域错误。', object_id=obj['id'])
    for p in doc.paragraphs:
        for f in fields(p):
            if ERROR_TEXT.search(f.get('cached', '')): bad('field-result-error', '更新域后的结果存在错误。', paragraph=doc.pindex[p])
        if p not in caps and doc.level(p) is None and not re.match(r'^(TOC|目录)', doc.style_name(p), re.I) and not p.xpath('ancestor::w:txbxContent', namespaces=NS):
            if next(body_references(visible(p)), None): bad('static-figure-table-reference', '正文仍有固定图表编号引用，需定位后改为上下图/表。', paragraph=doc.pindex[p])
    return {'status': 'passed' if not issues else 'failed', 'issues': issues,
            'template_sha256': contract['template_sha256'], 'input_sha256': file_digest(path),
            'object_count': len(objects), 'image_content_requires_visual_review': any(o['kind'] == 'figure' for o in objects),
            'field_update_required': True, 'note': '结构检查通过不代表已运行 Word F9，也不代表已目视检查图片。'}


def normalize_document(source: Path, out: Path, template: Path, style_json: Path | None = None, object_plan=None):
    source, out = Path(source), Path(out)
    if out.resolve() in {source.resolve(), Path(template).resolve()}: raise PolicyError('in-place-write', '输出不能覆盖输入或模板。')
    doc = Doc(source); plan = load_object_plan(object_plan, doc)
    contract = template_contract(template, required_levels(doc), style_json); changes = []
    normalize_headings(doc, contract, changes)
    old_refs = normalize_captions(doc, contract, plan, changes)
    relative_references(doc, old_refs, changes)
    settings = doc.root('word/settings.xml', 'settings'); set_child(settings, 'updateFields', val='true')
    doc.changed.add('word/settings.xml'); doc.register('settings')
    out.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(dir=out.parent) as temp:
        candidate = Path(temp) / 'candidate.docx'; doc.save(candidate)
        # Semantic choices were already source-bound. Auditing reuses object hashes,
        # not the now-stale original document hash.
        final_objects = {o['id']: o for o in inventory(Doc(candidate))}
        final_choices = [{**row, 'object_sha256': final_objects[ident]['hash']} for ident, row in plan.items()]
        audit_plan = {'objects': final_choices} if final_choices else None
        checked = audit_document(candidate, template, style_json, audit_plan)
        if checked['status'] != 'passed':
            raise PolicyError('numbering-policy-audit-failed', '新编号/题注规则复查未通过，未写入输出文件。', issues=checked['issues'])
        from layout_invariant_guard import snapshot_docx, INVARIANTS
        before, after = snapshot_docx(source), snapshot_docx(candidate)
        damaged = [k for k in INVARIANTS['automatic-numbering'] if before[k] != after[k]]
        if damaged: raise PolicyError('numbering-collateral-edit', '编号/题注修复改动了冻结对象。', invariants=damaged)
        os.replace(candidate, out)
    return {'status': 'passed', 'changes': changes, 'change_count': len(changes),
            'output': str(out), 'output_sha256': file_digest(out), 'audit': checked,
            'object_decisions': final_choices}


def main():
    parser = argparse.ArgumentParser(description='读取实际模板编号；补齐独立图表题注并使用章内自动序号')
    sub = parser.add_subparsers(dest='command', required=True)
    inv = sub.add_parser('inventory'); inv.add_argument('docx', type=Path); inv.add_argument('--json-out', type=Path)
    for action in ('repair', 'audit'):
        p = sub.add_parser(action); p.add_argument('docx', type=Path); p.add_argument('--template', type=Path, required=True)
        p.add_argument('--template-style-json', type=Path); p.add_argument('--object-plan', type=Path); p.add_argument('--json-out', type=Path)
        if action == 'repair': p.add_argument('--out', type=Path, required=True)
    args = parser.parse_args()
    try:
        if args.command == 'inventory':
            result = {'version': 1, 'source_sha256': file_digest(args.docx), 'objects': public_inventory(Doc(args.docx)),
                      'instruction': 'Agent 查看全部图表；缺题注时根据内容填写 title，非业务图表须提供 exempt/reason。不要把未检查项目填成通过。'}
        elif args.command == 'repair': result = normalize_document(args.docx, args.out, args.template, args.template_style_json, args.object_plan)
        else: result = audit_document(args.docx, args.template, args.template_style_json, args.object_plan)
    except (PolicyError, ValueError, OSError, zipfile.BadZipFile) as exc:
        result = {'status': 'blocked', 'issues': [exc.as_issue() if isinstance(exc, PolicyError) else {'code': 'numbering-policy-error', 'message': str(exc), 'severity': 'error'}]}
    encoded = json.dumps(result, ensure_ascii=False, indent=2)
    if args.json_out:
        args.json_out.parent.mkdir(parents=True, exist_ok=True); args.json_out.write_text(encoded + '\n', encoding='utf-8')
    print(encoded)
    return 2 if result.get('status') in {'failed', 'blocked'} else 0


if __name__ == '__main__':
    raise SystemExit(main())
