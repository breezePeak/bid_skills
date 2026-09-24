#!/usr/bin/env python3
"""Small, model-independent cursor and scope guard for one working DOCX.

This records local progress, not a visual PASS or a substitute for final review.
Pages are viewing windows; stable block IDs address complete body objects.
"""
from __future__ import annotations

import copy
import hashlib
import json
import os
import posixpath
import re
import shutil
import tempfile
import socket
import zipfile
from contextlib import contextmanager
from pathlib import Path
from lxml import etree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'w': W, 'r': R}
Q = lambda name: '{' + W + '}' + name
DOC = 'word/document.xml'
RELS = 'word/_rels/document.xml.rels'
STATE = 'block-state.json'


class BlockError(ValueError):
    pass


def sha(path):
    return 'sha256:' + hashlib.sha256(Path(path).read_bytes()).hexdigest()


def read_json(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


def write_json(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(dir=path.parent, suffix='.tmp')
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as out:
            json.dump(value, out, ensure_ascii=False, indent=2)
            out.write('\n')
        os.replace(name, path)
    finally:
        Path(name).unlink(missing_ok=True)


def parse(data):
    return ET.fromstring(data, ET.XMLParser(resolve_entities=False, no_network=True))


def signature(node):
    """Compare XML meaning without serialization/namespace-prefix noise."""
    if not isinstance(node.tag, str):
        return ('comment', node.text)
    text = node.text if node.tag in {Q('t'), Q('instrText'), Q('delText')} else (node.text or '').strip()
    volatile = {Q(k) for k in ('rsidR', 'rsidRPr', 'rsidDel', 'rsidRDefault', 'rsidP')}
    return (node.tag, tuple(sorted((k, v) for k, v in node.attrib.items() if k not in volatile)), text, tuple(signature(c) for c in node))


def same_part(a, b, name):
    if a == b:
        return True
    if a is None or b is None or not name.endswith(('.xml', '.rels')):
        return False
    try:
        return signature(parse(a)) == signature(parse(b))
    except ET.XMLSyntaxError:
        return False


def package(path):
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        if len(set(names)) != len(names) or archive.testzip():
            raise BlockError('DOCX 包损坏或存在重复部件。')
        if any(n.startswith('/') or '..' in n.split('/') for n in names):
            raise BlockError('DOCX 包含非法部件路径。')
        parts = {n: archive.read(n) for n in names}
    if DOC not in parts or 'word/styles.xml' not in parts:
        raise BlockError('不是完整 DOCX。')
    root = parse(parts[DOC])
    body = root.find(Q('body'))
    if body is None:
        raise BlockError('DOCX 缺少正文。')
    return parts, root, body


def visible(element):
    return ''.join(element.itertext()) if element.tag == Q('t') else ''.join(element.xpath('.//w:t/text()', namespaces=NS))


def make_blocks(path, target_chars=1800, max_paragraphs=14):
    """Conservative semantic grouping. The budgets are NOT claimed page counts."""
    if target_chars < 1 or max_paragraphs < 1:
        raise BlockError('分块预算必须为正数。')
    parts, root, body = package(path)
    nodes = list(body)
    styles = {s.get(Q('styleId')): s for s in parse(parts['word/styles.xml']).findall(Q('style'))}
    def style_heading(sid, seen=None):
        seen = set() if seen is None else seen
        if not sid or sid in seen or sid not in styles:
            return False
        seen.add(sid)
        s = styles[sid]
        name = s.find(Q('name'))
        name = name.get(Q('val'), '') if name is not None else ''
        if re.match(r'^(TOC|目录|Title$|Subtitle$|封面)', name, re.I):
            return False
        lv = s.find('w:pPr/w:outlineLvl', NS)
        if lv is not None:
            return lv.get(Q('val')) in set(map(str, range(9)))
        if re.fullmatch(r'(Heading|标题)\s*[1-9]', name, re.I):
            return True
        parent = s.find(Q('basedOn'))
        return style_heading(parent.get(Q('val')) if parent is not None else None, seen)
    def kind(n):
        if n.tag == Q('sectPr'):
            return 'section'
        if n.tag == Q('tbl'):
            return 'table'
        if n.tag != Q('p'):
            return 'container'  # SDT, alternate content etc. are indivisible.
        if n.xpath('.//w:drawing|.//w:pict|.//w:object', namespaces=NS):
            return 'figure'
        instr = ' '.join(n.xpath('.//w:instrText/text()|.//w:fldSimple/@w:instr', namespaces=NS))
        if re.search(r'\bSEQ\s+(图|表|Figure|Table)\b', instr, re.I) or re.match(r'^\s*(图|表|Figure|Table)\s*\d+', visible(n), re.I):
            return 'caption'
        lv = n.find('w:pPr/w:outlineLvl', NS)
        sid = n.find('w:pPr/w:pStyle', NS)
        if (lv is not None and lv.get(Q('val')) in set(map(str, range(9)))) or style_heading(sid.get(Q('val')) if sid is not None else None):
            return 'heading'
        return 'text'
    kinds = [kind(n) for n in nodes]
    def blank(k):
        return kinds[k] == 'text' and not visible(nodes[k]).strip()
    def skip_blank(k):
        end = k
        while end < len(nodes) and blank(end) and end - k < 2:
            end += 1
        return end
    groups = []
    i = 0
    while i < len(nodes):
        if kinds[i] == 'section':
            i += 1
            continue
        j = i + 1
        # A heading and the first real content unit are one block, not all following text.
        if kinds[i] == 'heading':
            while j < len(nodes) and (kinds[j] == 'heading' or blank(j)):
                j += 1
            if j < len(nodes) and kinds[j] != 'section':
                j += 1
        last = j - 1
        if kinds[last] in {'caption', 'table', 'figure'}:
            k = skip_blank(j)
            companion = (k < len(nodes) and
                         ((kinds[last] == 'caption' and kinds[k] in {'table', 'figure'}) or
                          (kinds[last] in {'table', 'figure'} and kinds[k] == 'caption')))
            if companion:
                j = k + 1
        types = set(kinds[i:j])
        groups.append([i, j, 'object' if types & {'table', 'figure', 'container'} else 'text'])
        i = j
    spans = []
    for a, b, typ in groups:
        if spans and typ == spans[-1][2] == 'text' and spans[-1][1] == a:
            lo = spans[-1][0]
            size = sum(len(visible(n)) for n in nodes[lo:b])
            if kinds[a] != 'heading' and b - lo <= max_paragraphs and size <= target_chars:
                spans[-1][1] = b
                continue
        spans.append([a, b, typ])
    blocks = [{'id': f'B{n:04d}', 'part': DOC, 'start': a, 'end': b,
               'source_start': a, 'source_end': b, 'kind': typ, 'status': 'pending',
               'note': '', 'preview': visible(nodes[a])[:100]}
              for n, (a, b, typ) in enumerate(spans, 1)]
    # Cover stories outside document.xml once; notes can be inspected alongside their
    # body references, but still have an explicit final coverage entry here.
    story_names = sorted(k for k in parts if re.fullmatch(
        r'word/(?:header[^/]*|footer[^/]*|footnotes|endnotes)\.xml', k))
    for name in story_names:
        story = parse(parts[name])
        if not len(story):
            continue
        blocks.append({'id': f'B{len(blocks)+1:04d}', 'part': name, 'kind': 'story',
                       'start': 0, 'end': len(story), 'source_start': 0,
                       'source_end': len(story), 'status': 'pending', 'note': '',
                       'preview': name + ': ' + visible(story)[:80]})
    return blocks


@contextmanager
def session_lock(work):
    """OS-held lock releases on process exit, including a crash; no stale mkdir lock."""
    path = Path(work) / '.block-session-v2.lock'
    path.parent.mkdir(parents=True, exist_ok=True)
    handle = open(path, 'a+b')
    acquired = False
    try:
        if os.name == 'nt':
            import msvcrt
            if handle.seek(0, os.SEEK_END) == 0:
                handle.write(b'0'); handle.flush()
            handle.seek(0)
            try:
                msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
            except OSError as exc:
                raise BlockError('另一个操作正在写入此会话。') from exc
        else:
            import fcntl
            try:
                fcntl.flock(handle, fcntl.LOCK_EX | fcntl.LOCK_NB)
            except BlockingIOError as exc:
                raise BlockError('另一个操作正在写入此会话。') from exc
        acquired = True
        yield
    finally:
        if acquired:
            if os.name == 'nt':
                handle.seek(0)
                msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(handle, fcntl.LOCK_UN)
        handle.close()


def _initialize_unlocked(source, work, settings=None, **budgets):
    source, work = Path(source).resolve(), Path(work).resolve()
    work.mkdir(parents=True, exist_ok=True)
    if (work / STATE).exists():
        state = load(work)
        if sha(source) != state['input_sha256']:
            raise BlockError('工作目录属于另一份输入；不能重置原文。')
        return state
    original = work / 'original.docx'
    if source == original:
        raise BlockError('原文保护路径冲突；请使用新的工作目录。')
    package(source)
    if original.exists():
        if sha(original) != sha(source):
            raise BlockError('遗留原文与输入不一致，不能覆盖。')
    else:
        shutil.copyfile(source, original)
    current = work / 'checkpoints' / 'r000000.docx'
    current.parent.mkdir(exist_ok=True)
    if current.exists():
        if sha(current) != sha(original):
            raise BlockError('遗留检查点与原文不一致，不能覆盖。')
    else:
        shutil.copyfile(original, current)
    state = {'version': 1, 'coverage_version': 2, 'workflow': 'block-v1', 'input_sha256': sha(source),
             'source': str(original), 'source_sha256': sha(original),
             'current': str(current), 'current_sha256': sha(current), 'revision': 0,
             'global_ready': False, 'settings': settings or {},
             'blocks': make_blocks(current, **budgets), 'boundary_dirty': [], 'history': [], 'final': None}
    write_json(work / STATE, state)
    return state


def initialize(source,work,settings=None,**budgets):
    Path(work).mkdir(parents=True,exist_ok=True)
    with session_lock(work):
        return _initialize_unlocked(source,work,settings,**budgets)


def load(work):
    state = read_json(Path(work) / STATE)
    if state.get('workflow') != 'block-v1' or state.get('version') != 1:
        raise BlockError('不是本版分块会话。')
    for name in ('source', 'current'):
        if not Path(state[name]).is_file() or sha(state[name]) != state[name + '_sha256']:
            raise BlockError(f'{name} 文件在检查点之外被改动；不得覆盖进度。')
    for name, ref in state['settings'].get('baselines', {}).items():
        if not Path(ref['path']).is_file() or sha(ref['path']) != ref['sha256']:
            raise BlockError(f'本轮 {name} 基准变了；不能沿用旧块结论。')
    if state.get('coverage_version',1) < 2:
        # Upgrade coverage in memory without resetting completed body blocks.
        # The next state-changing action persists it atomically under the session lock.
        parts,_,_ = package(state['current'])
        covered = {b.get('part',DOC) for b in state['blocks']}
        ids = {b['id'] for b in state['blocks']}
        for part in sorted(parts):
            if re.fullmatch(r'word/(?:header[^/]*|footer[^/]*|footnotes|endnotes)\.xml',part) and part not in covered:
                nodes = list(parse(parts[part]))
                if not nodes:
                    continue
                n = len(ids)+1
                while f'B{n:04d}' in ids:
                    n += 1
                ident = f'B{n:04d}'; ids.add(ident)
                state['blocks'].append({'id':ident,'part':part,'kind':'story','start':0,'end':len(nodes),'source_start':0,'source_end':len(nodes),'status':'pending','note':'补齐附属内容检查','preview':part})
        state['coverage_version']=2
        if state.get('final'):
            state['previous_final']=state['final']
        state['final']=None; state.pop('final_attempt',None)
    validate_blocks(state)
    return state


def validate_blocks(state):
    blocks = state.get('blocks')
    if not isinstance(blocks, list) or len({b.get('id') for b in blocks}) != len(blocks):
        raise BlockError('分块清单缺失或存在重复 ID。')
    parts, _, body = package(state['current'])
    coverage = set()
    for block in blocks:
        part = block.get('part', DOC)
        if part not in parts or block.get('status') not in {'pending', 'checked'}:
            raise BlockError('块对象丢失或状态无效。')
        nodes = list(body) if part == DOC else list(parse(parts[part]))
        a, b = block.get('start'), block.get('end')
        if type(a) is not int or type(b) is not int or not 0 <= a < b <= len(nodes):
            raise BlockError('块范围已失效，不能跳过或猜测位置。')
        for i in range(a,b):
            key = (part,i)
            if key in coverage:
                raise BlockError('块范围重叠。')
            coverage.add(key)
    expected = {(DOC,i) for i,n in enumerate(body) if n.tag != Q('sectPr')}
    # Auxiliary stories are required for new sessions; legacy sessions are upgraded on load.
    for name in parts:
        if re.fullmatch(r'word/(?:header[^/]*|footer[^/]*|footnotes|endnotes)\.xml', name):
            expected.update((name,i) for i in range(len(parse(parts[name]))))
    if not expected <= coverage:
        raise BlockError('分块清单有漏项；须重建缺少的附属内容块，不得提前验收。')


def block_by_id(state, ident):
    found = [b for b in state['blocks'] if b['id'] == ident]
    if len(found) != 1:
        raise BlockError('块 ID 不存在；请使用程序返回的 ID。')
    return found[0]


def next_block(work, view_offset=0, view_size=14):
    if view_offset < 0 or not 1 <= view_size <= 100:
        raise BlockError('查看窗口参数无效。')
    state = load(work)
    todo = next((b for b in state['blocks'] if b['status'] != 'checked'), None)
    result = {'current': state['current'], 'revision': state['revision'],
              'completed': sum(b['status'] == 'checked' for b in state['blocks']),
              'total': len(state['blocks']), 'global_ready': state['global_ready'],
              'status': 'global_setup' if not state['global_ready'] else 'block_review' if todo else 'ready_for_final',
              'block': copy.deepcopy(todo)}
    if todo:
        parts, _, body = package(state['current'])
        part = todo.get('part', DOC)
        nodes = list(body) if part == DOC else list(parse(parts[part]))
        units = []
        for n in nodes[todo['start']:todo['end']]:
            if n.tag == Q('tbl'):
                units.extend(' | '.join(visible(c) for c in row.findall(Q('tc'))) for row in n.findall(Q('tr')))
            else:
                units.append(visible(n))
        selected = units[view_offset:view_offset + view_size]
        result['block']['view'] = {'offset': view_offset, 'size': view_size, 'total_units': len(units),
                                  'has_more': view_offset + view_size < len(units),
                                  'text_truncated': any(len(t) > 1000 for t in selected)}
        result['block']['paragraphs'] = [t[:1000] for t in selected]
        note_refs = {kind: sorted(note_ids(nodes[todo['start']:todo['end']], kind)) for kind in ('footnote','endnote')}
        result['block']['note_references'] = note_refs
        result['block']['context_before'] = visible(nodes[todo['start'] - 1])[-400:] if todo['start'] else ''
        result['block']['context_after'] = visible(nodes[todo['end']])[:400] if todo['end'] < len(nodes) else ''
        folder = Path(work) / 'block-data'; folder.mkdir(exist_ok=True)
        stem = todo['id'] + '-r' + str(state['revision'])
        fragment = folder / (stem + '.xml'); textfile = folder / (stem + '.txt')
        fragment.write_text(''.join(ET.tostring(n, encoding='unicode') for n in nodes[todo['start']:todo['end']]), encoding='utf-8')
        textfile.write_text('\n'.join(units), encoding='utf-8')
        result['block']['xml_path'] = str(fragment.resolve())
        result['block']['text_path'] = str(textfile.resolve())
    return result


def relmap(parts, name=RELS):
    return {} if name not in parts else {n.get('Id'): dict(n.attrib) for n in parse(parts[name])}


def refs(nodes):
    return {v for n in nodes for e in n.iter() for k, v in e.attrib.items() if k.startswith('{' + R + '}')}


def target(rel, owner='word/document.xml'):
    if rel.get('TargetMode') == 'External':
        return None
    t = rel.get('Target', '')
    return posixpath.normpath(t.lstrip('/') if t.startswith('/') else posixpath.join(posixpath.dirname(owner), t))


def note_ids(nodes, kind):
    return {e.get(Q('id')) for n in nodes for e in n.iter(Q(kind + 'Reference'))}


def note_scope(old, new, part, before_nodes, after_nodes, outside_nodes):
    kind = 'footnote' if part.endswith('footnotes.xml') else 'endnote'
    before_ids, after_ids = note_ids(before_nodes, kind), note_ids(after_nodes, kind)
    protected = note_ids(outside_nodes, kind)
    oldroot = parse(old[part]) if part in old else None
    newroot = parse(new[part]) if part in new else None
    oldnotes = {} if oldroot is None else {n.get(Q('id')): signature(n) for n in oldroot}
    newnotes = {} if newroot is None else {n.get(Q('id')): signature(n) for n in newroot}
    owned = (before_ids | after_ids) - protected
    if any(newnotes.get(k) != v for k, v in oldnotes.items() if k not in owned):
        raise BlockError('修改影响了当前块以外的脚注/尾注。')
    for n in ([] if newroot is None else newroot):
        ident = n.get(Q('id'))
        separator = oldroot is None and n.get(Q('type')) in {'separator', 'continuationSeparator'}
        if ident not in oldnotes and ident not in after_ids and not separator:
            raise BlockError('新增脚注/尾注没有绑定当前块。')


def append_only_definitions(old, new, part, outside_nodes):
    """New local IDs may be installed, but existing shared definitions cannot change."""
    if part not in new:
        return False
    b = parse(new[part])
    a = parse(old[part]) if part in old else ET.Element(b.tag, dict(b.attrib))
    if a.tag != b.tag or a.attrib != b.attrib:
        return False
    def key(n):
        if n.tag == Q('style'):
            return (n.tag, n.get(Q('styleId')))
        if n.tag == Q('num'):
            return (n.tag, n.get(Q('numId')))
        if n.tag == Q('abstractNum'):
            return (n.tag, n.get(Q('abstractNumId')))
        return (n.tag, None)
    aa, bb = {key(n): signature(n) for n in a}, {key(n): signature(n) for n in b}
    if len(aa) != len(a) or len(bb) != len(b) or any(bb.get(k) != v for k, v in aa.items()):
        return False
    for n in b:
        if key(n) in aa:
            continue
        if n.tag == Q('style'):
            ident = n.get(Q('styleId'))
            if n.get(Q('default')) in {'1', 'true', 'on'}:
                return False
            if any(e.get(Q('val')) == ident for e in a.iter() if e.tag in {Q('basedOn'), Q('link'), Q('next')}):
                return False
            if any(e.get(Q('val')) == ident for out in outside_nodes for e in out.iter() if e.tag in {Q('pStyle'), Q('rStyle'), Q('tblStyle')}):
                return False
        elif n.tag == Q('num'):
            if any(e.get(Q('val')) == n.get(Q('numId')) for out in outside_nodes for e in out.iter(Q('numId'))):
                return False
        elif n.tag == Q('abstractNum'):
            if any(e.get(Q('val')) == n.get(Q('abstractNumId')) for e in a.iter(Q('abstractNumId'))):
                return False
        else:
            return False
    return True


def local_scope(before, after, block):
    """Allow only the selected body range and its exclusively-owned image assets."""
    if block.get('part', DOC) != DOC:
        return story_scope(before, after, block)
    old, oroot, obody = package(before)
    new, nroot, nbody = package(after)
    a, b = block['start'], block['end']
    on, nn = list(obody), list(nbody)
    if oroot.attrib != nroot.attrib or obody.attrib != nbody.attrib or [signature(n) for n in oroot if n is not obody] != [signature(n) for n in nroot if n is not nbody]:
        raise BlockError('局部修复改动了正文之外的文档设置。')
    tail = len(on) - b
    end = len(nn) - tail
    if end <= a or [signature(n) for n in on[:a]] != [signature(n) for n in nn[:a]] or [signature(n) for n in on[b:]] != [signature(n) for n in nn[end:]]:
        raise BlockError('修改超出当前块或删除了块边界；候选未采纳。')
    # Section definitions are shared layout, even when nested inside a paragraph.
    sections = lambda ns: [signature(s) for n in ns for s in n.iter(Q('sectPr'))]
    if sections(on[a:b]) != sections(nn[a:end]):
        raise BlockError('局部修复不能改变分节/页面设置。')
    oldrels, newrels = relmap(old), relmap(new)
    owned = refs(on[a:b]); outside = refs(on[:a] + on[b:])
    used_new = refs(nn[a:end])
    for rid in set(oldrels) | set(newrels):
        if oldrels.get(rid) == newrels.get(rid):
            continue
        row = newrels.get(rid)
        note_part = next((kind for kind in ('footnotes', 'endnotes') if row and row.get('Type', '').endswith('/' + kind)), None)
        if note_part and rid not in oldrels and target(row) == 'word/' + note_part + '.xml' and note_ids(nn[a:end], note_part[:-1]):
            continue
        if rid in outside or row is None or rid not in used_new or not row.get('Type', '').endswith('/image') or row.get('TargetMode') == 'External':
            raise BlockError('修改影响块外关系，或新增了未绑定的关系。')
        if rid in oldrels and (rid not in owned or not oldrels[rid].get('Type', '').endswith('/image')):
            raise BlockError('不能把现有非图片关系改作图片。')
    protected_media = {target(oldrels[r]) for r in outside if r in oldrels}
    for name in old:
        if name.endswith('.rels') and name != RELS:
            owner = name.replace('/_rels/', '/').removesuffix('.rels')
            protected_media.update(target(r, owner) for r in relmap(old, name).values())
    for rid in used_new:
        rel = newrels.get(rid)
        if rel and rel.get('Type', '').endswith('/image') and rel.get('TargetMode') != 'External' and target(rel) not in new:
            raise BlockError('当前块图片关系指向不存在的媒体。')
    outside_definitions = on[:a] + on[b:] + [parse(v) for k, v in old.items() if k.endswith('.xml') and k != DOC and (k in {'word/styles.xml', 'word/numbering.xml', 'word/footnotes.xml', 'word/endnotes.xml'} or re.match(r'word/(header|footer)\d+\.xml$', k))]
    allowed_media = {target(newrels[r]) for r in used_new if r in newrels} | {target(oldrels[r]) for r in owned if r in oldrels}
    changes = []
    for name in set(old) | set(new):
        if same_part(old.get(name), new.get(name), name):
            continue
        changes.append(name)
        if name in {DOC, RELS}:
            continue
        if name.startswith('word/media/') and name in allowed_media and name not in protected_media:
            continue
        if name in {'word/footnotes.xml', 'word/endnotes.xml'}:
            note_scope(old, new, name, on[a:b], nn[a:end], on[:a]+on[b:])
            continue
        if name in {'word/styles.xml', 'word/numbering.xml'} and append_only_definitions(old, new, name, outside_definitions):
            continue
        if name == '[Content_Types].xml':
            o = {signature(n) for n in parse(old[name])}
            nr = list(parse(new[name]))
            added = [n for n in nr if signature(n) not in o]
            if o <= {signature(n) for n in nr} and all(n.get('ContentType', '').startswith('image/') or (n.get('PartName') in {'/word/footnotes.xml', '/word/endnotes.xml'} and n.get('ContentType') == 'application/vnd.openxmlformats-officedocument.wordprocessingml.' + n.get('PartName').rsplit('/', 1)[-1][:-4] + '+xml' and n.get('PartName')[1:] in new) for n in added):
                continue
        raise BlockError(f'局部修复修改了共享或块外部件 {name}；候选未采纳。')
    return end - b, sorted(changes)


def global_scope(before, after):
    old, _, obody = package(before)
    new, _, nbody = package(after)
    def without_sections(body):
        node = copy.deepcopy(body)
        for part in list(node.iter(Q('sectPr'))):
            part.getparent().remove(part)
        # Empty pPr left after removing a section has no visible effect.
        for pp in list(node.iter(Q('pPr'))):
            if len(pp) == 0 and not pp.attrib:
                pp.getparent().remove(pp)
        return signature(node)
    if len(obody) != len(nbody) or without_sections(obody) != without_sections(nbody):
        raise BlockError('全局准备只设置共享样式、页面和页眉页脚，不批量改正文。')
    allowed = {DOC, RELS, '[Content_Types].xml', 'word/styles.xml', 'word/stylesWithEffects.xml',
               'word/numbering.xml', 'word/settings.xml', 'word/fontTable.xml'}
    for name in set(old) | set(new):
        if same_part(old.get(name), new.get(name), name):
            continue
        if name in allowed or re.match(r'word/(theme/|header\d+\.xml$|footer\d+\.xml$|_rels/(header|footer)\d+\.xml\.rels$)', name):
            continue
        raise BlockError(f'全局准备不能修改业务内容部件：{name}')
    # Existing body image/hyperlink relationships may not be redirected by setup.
    orels, nrels = relmap(old), relmap(new)
    content = copy.deepcopy(obody)
    for section in list(content.iter(Q('sectPr'))):
        section.getparent().remove(section)
    for rid in refs(list(content)):
        if orels.get(rid) != nrels.get(rid):
            raise BlockError('全局准备改变了正文对象的关系。')


def _save_revision(work, state, proposal, event):
    work = Path(work)
    state['revision'] += 1
    proposal = Path(proposal)
    unchanged = sha(proposal) == state['current_sha256']
    out = Path(state['current']) if unchanged else work / 'checkpoints' / f"r{state['revision']:06d}.docx"
    if out.exists():
        if sha(out) != sha(proposal):
            raise BlockError('存在未完成的不同内容检查点；请先核对，不覆盖它。')
    else:
        shutil.copyfile(proposal, out)
    state['current'], state['current_sha256'] = str(out.resolve()), sha(out)
    parts,_,_ = package(out)
    covered = {b.get('part',DOC) for b in state['blocks']}
    for part in sorted(parts):
        if re.fullmatch(r'word/(?:header[^/]*|footer[^/]*|footnotes|endnotes)\.xml',part) and part not in covered:
            root = parse(parts[part])
            if len(root):
                state['blocks'].append({'id':f"B{len(state['blocks'])+1:04d}",'part':part,'kind':'story','start':0,'end':len(root),'source_start':0,'source_end':0,'status':'pending','note':'新建附属内容须检查','preview':part})
    if state.get('final'):
        state['previous_final'] = state['final']
    state['final'] = None
    state.pop('final_attempt', None)
    state['history'].append({'revision': state['revision'], **event})
    write_json(work / STATE, state)
    return state


def accept_global(work, proposal, note):
    if not note.strip():
        raise BlockError('请记录实际全局检查/修改结果。')
    with session_lock(work):
        state = load(work)
        if state['global_ready'] or any(b['status'] == 'checked' for b in state['blocks']):
            raise BlockError('全局准备已经结束，不能重新批量处理全文。')
        global_scope(state['current'], proposal)
        state['global_ready'] = True
        return _save_revision(work, state, proposal, {'action': 'global', 'note': note})


def checkpoint(work, ident, proposal, note):
    if not note.strip():
        raise BlockError('请记录当前块实际检查了什么、改了什么。')
    with session_lock(work):
        state = load(work)
        if not state['global_ready']:
            raise BlockError('先完成一次全局模板/页面设置。')
        block = block_by_id(state, ident)
        if block['status'] == 'checked':
            raise BlockError('该块已经完成；发现新问题时先定点 reopen，不能无故重做。')
        first = next(b for b in state['blocks'] if b['status'] != 'checked')
        if first['id'] != ident:
            raise BlockError('请先处理当前待修块；不得跳过前面的块。')
        delta, changes = local_scope(state['current'], proposal, block)
        block['end'] += delta
        at = state['blocks'].index(block)
        for other in state['blocks'][at + 1:]:
            if other.get('part', DOC) == block.get('part', DOC):
                other['start'] += delta
                other['end'] += delta
        block.update(status='checked', note=note)
        proposal_parts, _, _ = package(proposal)
        covered = {b.get('part', DOC) for b in state['blocks']}
        for part in sorted(proposal_parts):
            if re.fullmatch(r'word/(?:header[^/]*|footer[^/]*|footnotes|endnotes)\.xml', part) and part not in covered:
                root = parse(proposal_parts[part])
                if len(root):
                    state['blocks'].append({'id': f"B{len(state['blocks'])+1:04d}", 'part':part, 'kind':'story', 'start':0, 'end':len(root), 'source_start':0, 'source_end':0, 'status':'pending', 'note':'新增附属内容需要检查', 'preview':part})
        for other in state['blocks']:
            part = other.get('part', DOC)
            if other['kind'] == 'story' and part in proposal_parts:
                other['end'] = len(parse(proposal_parts[part]))
                if part in changes and other is not block:
                    other['status'] = 'pending'
        if changes:
            state['boundary_dirty'] = sorted(set(state['boundary_dirty']) | {b['id'] for b in state['blocks'][max(0, at-1):at+2]})
        return _save_revision(work, state, proposal, {'action': 'block', 'block': ident, 'note': note, 'parts_changed': changes})


def reopen(work, ident, reason):
    if not reason.strip():
        raise BlockError('返回修复需要具体问题，不能无故重跑。')
    with session_lock(work):
        state = load(work)
        block = block_by_id(state, ident)
        block.update(status='pending', note=reason)
        if state.get('final'):
            state['previous_final'] = state['final']
        state['final'] = None
        state.pop('final_attempt',None)
        state['history'].append({'revision': state['revision'], 'action': 'reopen', 'block': ident, 'note': reason})
        write_json(Path(work) / STATE, state)
        return state


def validate_final_binding(manifest):
    """Called by the existing release gate; a local checkpoint is not release."""
    ref = manifest.get('block_session')
    if not isinstance(ref, dict) or not Path(ref.get('path', '')).is_file() or sha(ref['path']) != ref.get('sha256'):
        raise BlockError('缺少当前分块进度，或验收之后进度已改变。')
    state = load(Path(ref['path']).parent)
    if not state['global_ready'] or any(b['status'] != 'checked' for b in state['blocks']):
        raise BlockError('还有未完成块，不能交付。')
    final = state.get('final') or {}
    if final.get('candidate_sha256') != manifest.get('candidate_sha256') or final.get('revision') != state['revision']:
        raise BlockError('终检结果不属于最后一轮分块修复。')
    return state


def story_scope(before, after, block):
    """One complete header/footer/notes part, with exclusively-owned media."""
    old, _, _ = package(before)
    new, _, _ = package(after)
    part = block['part']
    if part not in new:
        raise BlockError('不能删除附属内容块。')
    relpart = posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')
    oldrels, newrels = relmap(old, relpart), relmap(new, relpart)
    current_ids = refs([parse(new[part])])
    old_ids = refs([parse(old[part])])
    protected = set()
    for name in old:
        if name.endswith('.rels') and name != relpart:
            owner = name.replace('/_rels/', '/').removesuffix('.rels')
            protected.update(target(r, owner) for r in relmap(old, name).values())
    allowed = {target(r, part) for r in list(oldrels.values()) + list(newrels.values())
               if r.get('Type', '').endswith('/image') and r.get('TargetMode') != 'External'}
    for rid in set(oldrels) | set(newrels):
        if oldrels.get(rid) == newrels.get(rid):
            continue
        rel = newrels.get(rid)
        if not rel or rid not in current_ids or not rel.get('Type','').endswith('/image') or rel.get('TargetMode') == 'External':
            raise BlockError('附属块修改了非图片关系。')
    for rid in current_ids:
        rel = newrels.get(rid)
        if rel and rel.get('Type','').endswith('/image') and rel.get('TargetMode') != 'External' and target(rel,part) not in new:
            raise BlockError('附属块图片关系失效。')
    changed = []
    outside = [parse(v) for k,v in old.items() if k.endswith('.xml') and k != part]
    for name in set(old) | set(new):
        if same_part(old.get(name), new.get(name), name):
            continue
        changed.append(name)
        if name in {part, relpart}:
            continue
        if name.startswith('word/media/') and name in allowed and name not in protected:
            continue
        if name in {'word/styles.xml','word/numbering.xml'} and append_only_definitions(old,new,name,outside):
            continue
        if name == '[Content_Types].xml':
            a = {signature(n) for n in parse(old[name])}; rows = list(parse(new[name]))
            if a <= {signature(n) for n in rows} and all(n.get('ContentType','').startswith('image/') for n in rows if signature(n) not in a):
                continue
        raise BlockError('附属块修改影响共享或块外部件：' + name)
    return len(parse(new[part])) - len(parse(old[part])), sorted(changed)


def accept_shared_repair(work, proposal, reason):
    """Explicit late shared-format fix: retain content progress; invalidate final QA."""
    if not reason.strip():
        raise BlockError('修复共享设置必须说明实际缺陷和影响范围。')
    with session_lock(work):
        state = load(work)
        if not state['global_ready']:
            raise BlockError('尚未完成全局准备。')
        global_scope(state['current'], proposal)
        state['boundary_dirty'] = [b['id'] for b in state['blocks']]
        return _save_revision(work,state,proposal,{'action':'shared-repair','note':reason})
