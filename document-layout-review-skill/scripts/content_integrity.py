#!/usr/bin/env python3
"""Source-to-final business-text guard, independent of a stored PASS.

Only explicitly recognized layout transformations are canonicalized. In
particular, ordinary business numbers are never stripped. Unrecognized content
changes block delivery, including during --audit-only continuation.
"""
from __future__ import annotations

import argparse
import difflib
import hashlib
import json
import re
import zipfile
from collections import Counter
from pathlib import Path
from lxml import etree as E

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
NS = {'w': W}
Q = lambda tag: '{' + W + '}' + tag
CJK = '\u3400-\u9fff\uf900-\ufaff'
DERIVED = {'SEQ', 'PAGE', 'NUMPAGES', 'SECTIONPAGES', 'PAGEREF', 'STYLEREF', 'TOC'}
CAPTION = re.compile(r'^\s*(图|表|Figure|Table)\s*(\d+(?:[.．\-－]\d+)*)\s+(.+)$', re.I)
HEAD_PREFIX = re.compile(r'^\s*(?:第[零〇一二三四五六七八九十百千万两\d]+[章节篇部]\s*|\d+(?:[.．]\d+)*(?:[.．、])?\s+)(?=\S)')
REFERENCE = re.compile(r'(?<![A-Za-z0-9_/\\.\-])(?:图|Figure)\s*\d+(?:[.．\-－]\d+)*|(?<![A-Za-z0-9_/\\.\-])(?:表|Table)\s*\d+(?:[.．\-－]\d+)*', re.I)


class ContentIntegrityError(ValueError):
    def __init__(self, code, message, **details):
        super().__init__(message)
        self.code, self.details = code, details

    def as_issue(self):
        return {'severity': 'error', 'code': self.code, 'message': str(self), **self.details}


def digest(path):
    return 'sha256:' + hashlib.sha256(Path(path).read_bytes()).hexdigest()


def parse(data):
    return E.fromstring(data, E.XMLParser(resolve_entities=False, no_network=True))


def read_json(value):
    if value is None:
        return None
    if isinstance(value, dict):
        return value
    return json.loads(Path(value).read_text(encoding='utf-8-sig'))


def canonical_text(value, *, references=False):
    # Preserve Latin spaces, signs, decimal points and all business digits.
    value = re.sub('[\u200b\ufeff]', '', value).replace('\u00a0', ' ').replace('\u3000', ' ')
    value = value.strip()
    for _ in range(2):
        value = re.sub(f'([{CJK}])\\s+(?=[{CJK}])', r'\1', value)
        value = re.sub(f'([{CJK}])\\s+(?=[，。；：！？、）】”’])', r'\1', value)
        value = re.sub(f'([（【“‘])\\s+(?=[{CJK}])', r'\1', value)
        value = re.sub(r'(\d)\s+(?=(?:年|月|日|时|分|秒|套|份|页|个|项|人|天|台|次|米|万元|元)(?![A-Za-z]))', r'\1', value)
        value = re.sub(r'([年月日时分秒])\s+(?=\d)', r'\1', value)
        value = re.sub(f'([{CJK}])\\s+(?=\\d+\\s*(?:套|份|页|个|项|人|天|台|次|米|元))', r'\1', value)
    # Canonicalize full/half-width punctuation only in Chinese prose.
    if re.search(f'[{CJK}]', value):
        value = value.translate(str.maketrans({'，': ',', '；': ';', '：': ':', '！': '!', '？': '?',
                                              '（': '(', '）': ')', '“': '"', '”': '"', '‘': "'", '’': "'"}))
        value = re.sub(f'(?<=[{CJK}])[。.]', '。', value)
        value = re.sub(r'\s+([,;:!?)])', r'\1', value)
        value = re.sub(r'([(])\s+', r'\1', value)
    if references:
        def marker(match):
            return '<figure-ref>' if re.match(r'图|Figure', match.group(), re.I) else '<table-ref>'
        value = REFERENCE.sub(marker, value)
        value = re.sub(r'(?:上|下)图', '<figure-ref>', value)
        value = re.sub(r'(?:上|下)表', '<table-ref>', value)
    return value


def styles_index(z):
    root = parse(z.read('word/styles.xml'))
    styles = {s.get(Q('styleId')): s for s in root.findall(Q('style'))}
    def heading(sid):
        seen = set()
        while sid and sid not in seen:
            seen.add(sid)
            s = styles.get(sid)
            if s is None:
                break
            level = s.find('w:pPr/w:outlineLvl', NS)
            if level is not None:
                try:
                    return 0 <= int(level.get(Q('val'), '9')) <= 8
                except ValueError:
                    return False
            name = s.find(Q('name'))
            if name is not None and re.fullmatch(r'(?:heading|标题)\s*[1-9]', name.get(Q('val'), ''), re.I):
                return True
            base = s.find(Q('basedOn'))
            sid = base.get(Q('val')) if base is not None else None
        return False
    return styles, heading



def protected_literals(text):
    # Literal addresses/code cannot be normalized as Chinese punctuation.
    return re.findall(r'`[^`]+`|"[A-Za-z]:[\\/][^"]*"|(?:https?://|ftp://|www\.)[^\s<>"\']+|(?<!\w)[A-Za-z]:[\\/][^\s<>"\']+',text)


def math_signature(node):
    # Formatting-only children do not change formula content/structure.
    excluded={Q('rPr'),'{'+M+'}rPr','{'+M+'}ctrlPr'}
    def shape(n):
        return [n.tag,sorted(n.attrib.items()),n.text or '',[shape(c) for c in n if c.tag not in excluded]]
    return hashlib.sha256(json.dumps(shape(node),ensure_ascii=False).encode()).hexdigest()

def visible_records(root):
    """Track complex fields across paragraphs; never double-count textbox text."""
    records, stack = [], []
    by_element = {}
    def visit(n, owner=None, omit=False):
        if not isinstance(n.tag, str):
            return
        if n.tag in {Q('drawing'), Q('pict'), Q('object'), Q('txbxContent')}:
            return  # image/Shape content is independently guarded and visually reviewed
        if n.tag in {Q('del'), Q('moveFrom')}:
            # Tracked-deletion text is protected separately, rather than silently lost.
            omit = True
        if n.tag == Q('p'):
            owner = {'node': n, 'text': [], 'fields': [], 'deleted': [], 'index': len(records) + 1}
            records.append(owner)
            by_element[n] = owner
        if n.tag=='{'+M+'}oMath' and owner is not None:
            owner['text'].append('<math:'+math_signature(n)+'>')
            return
        if n.tag == Q('fldSimple'):
            instruction = n.get(Q('instr'), '').strip()
            field = {'instruction': instruction, 'separate': True}
            stack.append(field)
            if owner is not None:
                owner['fields'].append(instruction)
            for c in n:
                visit(c, owner, omit)
            stack.pop()
            return
        if n.tag == Q('fldChar'):
            kind = n.get(Q('fldCharType'))
            if kind == 'begin':
                stack.append({'instruction': '', 'separate': False})
            elif kind == 'separate' and stack:
                stack[-1]['separate'] = True
                if owner is not None:
                    owner['fields'].append(stack[-1]['instruction'].strip())
            elif kind == 'end' and stack:
                stack.pop()
        elif n.tag == Q('instrText') and stack:
            stack[-1]['instruction'] += n.text or ''
        elif n.tag in {Q('t'), Q('delText'), Q('tab'), Q('br'), Q('cr'), Q('sym'), Q('footnoteReference'), Q('endnoteReference')} and owner is not None:
            derived = any(f['instruction'].strip().split(None, 1)[0].upper() in DERIVED
                          for f in stack if f['instruction'].strip())
            if derived:
                return
            text = n.text or '' if n.tag in {Q('t'), Q('delText')} else ('\t' if n.tag == Q('tab') else '\n')
            if n.tag in {Q('footnoteReference'),Q('endnoteReference')}:
                text='<'+E.QName(n).localname+':'+n.get(Q('id'),'')+'>'
            if n.tag == Q('sym'):
                text = '<sym:' + n.get(Q('font'), '') + ':' + n.get(Q('char'), '') + '>'
            owner['deleted' if omit else 'text'].append(text)
        for c in n:
            visit(c, owner, omit)
    visit(root)
    if stack:
        raise ContentIntegrityError('content-field-unclosed', '原文或候选存在未闭合域，不能完成内容核对。')
    return records


def is_object(node, kind):
    if kind == 'table':
        return node.tag == Q('tbl')
    return node.tag == Q('p') and bool(node.xpath('.//w:drawing|.//w:pict|.//w:object', namespaces=NS))


def caption_kind(record, styles):
    p, text = record['node'], ''.join(record['text']).strip()
    seq = [f for f in record['fields'] if re.match(r'^SEQ\s+(图|表|Figure|Table)(?:\s|$)', f, re.I)]
    if seq:
        return 'figure' if re.match(r'^SEQ\s+(图|Figure)(?:\s|$)', seq[0], re.I) else 'table'
    match = CAPTION.match(text)
    if not match:
        return None
    kind = 'figure' if match.group(1).lower() in {'图', 'figure'} else 'table'
    style = p.find('w:pPr/w:pStyle', NS)
    sid = style.get(Q('val')) if style is not None else None
    s = styles.get(sid)
    name = s.find(Q('name')) if s is not None else None
    if name is not None and re.search(r'caption|题注|图注|表注', name.get(Q('val'), ''), re.I):
        return kind
    # A manual-looking label alone is insufficient; require adjacent real object.
    for forward in (False, True):
        sibling=p.getnext() if forward else p.getprevious()
        for _ in range(3):
            if sibling is None:break
            if is_object(sibling,kind):return kind
            if sibling.tag!=Q('p') or ''.join(sibling.itertext()).strip():break
            sibling=sibling.getnext() if forward else sibling.getprevious()
    return None


def caption_title(record, kind):
    text = ''.join(record['text']).strip()
    match = CAPTION.match(text)
    if match:
        return match.group(3)
    label = r'(?:图|Figure)' if kind == 'figure' else r'(?:表|Table)'
    # The SEQ cache has been omitted, not the business title following it.
    return re.sub(r'^\s*' + label + r'\s*', '', text, count=1, flags=re.I)



def expand_vertical_merges(root):
    """Expand only the in-memory comparison grid; never edit the delivered file.

    Repeating a merged anchor on its logical rows makes a legitimate merge of
    equal labels compare equal. Different labels, orphan continuations and
    inconsistent spans cannot be hidden by declaring a cell merged.
    """
    import copy
    for table in reversed(list(root.iter(Q('tbl')))):
        active={}
        for row in table.findall(Q('tr')):
            before=row.find('w:trPr/w:gridBefore',NS)
            col=int(before.get(Q('val'),'0')) if before is not None else 0
            next_active={}
            for cell in row.findall(Q('tc')):
                span=cell.find('w:tcPr/w:gridSpan',NS)
                width=int(span.get(Q('val'),'1')) if span is not None else 1
                if width<1:raise ContentIntegrityError('content-grid-span-invalid','表格网格跨度无效。')
                cols=tuple(range(col,col+width));col+=width
                merged=cell.find('w:tcPr/w:vMerge',NS)
                if merged is None:continue
                if merged.get(Q('val'),'continue')=='restart':
                    anchor=(cell,cols)
                else:
                    anchor=active.get(cols[0])
                    if anchor is None or anchor[1]!=cols or any(active.get(c)!=anchor for c in cols):
                        raise ContentIntegrityError('content-merge-orphan','纵向合并缺少对应起始单元格或列跨度变化。')
                    base=anchor[0]
                    own=''.join(t.text or '' for t in cell.iter(Q('t')))
                    source=''.join(t.text or '' for t in base.iter(Q('t')))
                    if own.strip() and canonical_text(own)!=canonical_text(source):
                        raise ContentIntegrityError('content-merge-hidden-text','合并续接单元格含不同内容，不能隐藏或丢弃。')
                    if cell.xpath('.//w:drawing|.//w:pict|.//w:object|.//w:footnoteReference',namespaces=NS):
                        raise ContentIntegrityError('content-merge-hidden-object','合并续接单元格含独立对象，须先核对。')
                    for child in list(cell):
                        if child.tag!=Q('tcPr'):cell.remove(child)
                    for child in base:
                        if child.tag!=Q('tcPr'):cell.append(copy.deepcopy(child))
                for c in cols:next_active[c]=anchor
            active=next_active

def project(path, movements=()):
    result = {'body': [], 'captions': {'figure': [], 'table': []}, 'other_stories': {}, 'tracked_deletions': [], 'protected_literals': []}
    with zipfile.ZipFile(path) as z:
        styles, inherited_heading = styles_index(z)
        for part in sorted(z.namelist()):
            if not part.startswith('word/') or not part.endswith('.xml'):
                continue
            if not Path(part).stem.startswith(('document', 'header', 'footer', 'footnotes', 'endnotes', 'comments')):
                continue
            root = parse(z.read(part))
            original_indices={r['node']:r['index'] for r in visible_records(root)}
            expand_vertical_merges(root)
            for rec in visible_records(root):
                p = rec['node']
                original_index=original_indices.get(p)
                raw = ''.join(rec['text'])
                if part=='word/document.xml':
                    if any(m['note_paragraph']==original_index for m in movements):continue
                    edits=[m for m in movements if m['reference_paragraph']==original_index]
                    for edit in sorted(edits,key=lambda m:m['start'],reverse=True):
                        if raw[edit['start']:edit['end']]!=edit['marker']:
                            raise ContentIntegrityError('content-footnote-plan-stale','脚注引用不匹配原始定位。')
                        raw=raw[:edit['start']]+'<footnoteReference:'+edit['note_id']+'>'+raw[edit['end']:]
                deleted = ''.join(rec['deleted'])
                if deleted:
                    result['tracked_deletions'].append((part, deleted))
                if not raw.strip():
                    continue
                kind = caption_kind(rec, styles) if part == 'word/document.xml' else None
                if kind:
                    result['captions'][kind].append(canonical_text(caption_title(rec, kind)))
                    continue
                style = p.find('w:pPr/w:pStyle', NS)
                sid = style.get(Q('val')) if style is not None else None
                own_level = p.find('w:pPr/w:outlineLvl', NS)
                is_heading = (own_level is not None and own_level.get(Q('val')) in {str(i) for i in range(9)}) or inherited_heading(sid)
                in_cell = bool(p.xpath('ancestor::w:tc', namespaces=NS))
                if is_heading and not in_cell and not re.match(r'^\s*(?:19|20)\d{2}\s*(?:年|年度|[.．/\-]\d)',raw):
                    raw = HEAD_PREFIX.sub('', raw, count=1)
                result['protected_literals'].extend(protected_literals(raw))
                text = canonical_text(raw, references=(part == 'word/document.xml' and not in_cell and not is_heading))
                if part == 'word/document.xml':
                    # Keep paragraph and table-cell boundaries, not just concatenated text.
                    result['body'].append(('cell' if in_cell else 'paragraph', text))
                else:
                    note=next((a for a in p.iterancestors() if a.tag in {Q('footnote'),Q('endnote')}),None)
                    key=part+('#'+note.get(Q('id'),'')) if note is not None else part
                    result['other_stories'].setdefault(key, []).append(text)
    for move in movements:
        result['other_stories']['word/footnotes.xml#'+move['note_id']]=[canonical_text(move['note_body'])]
        result['protected_literals'].extend(protected_literals(move['note_body']))
    result['protected_literals'].sort()
    return result


def make_plan(source, object_plan=None):
    data = read_json(object_plan)
    captions = []
    if data:
        if data.get('source_sha256') != digest(source):
            raise ContentIntegrityError('content-plan-stale', '题注计划不属于原始文档，不能登记允许的新增文字。')
        for row in data.get('objects', []):
            # Existing-caption adjustments are not authorization to invent new business text.
            title = row.get('title', row.get('caption_title'))
            if title is not None:
                if not isinstance(title, str) or not title.strip() or not row.get('id') or not row.get('object_sha256'):
                    raise ContentIntegrityError('content-plan-invalid', '新增题注必须有真实对象定位、哈希和确定题名。')
                kind = 'figure' if str(row['id']).startswith(('figure', 'fig')) or re.fullmatch(r'F\d{4}',str(row['id'])) else 'table' if str(row['id']).startswith(('table', 'tbl')) or re.fullmatch(r'T\d{4}',str(row['id'])) else row.get('kind')
                if kind not in {'figure', 'table'}:
                    raise ContentIntegrityError('content-plan-invalid', '新增题注对象类型不确定。')
                captions.append({'id': row['id'], 'object_sha256': row['object_sha256'],
                                 'kind': kind, 'title': title})
    return {'version': 1, 'source_sha256': digest(source), 'allowed_caption_insertions': captions}



def _paragraph_records(source):
    with zipfile.ZipFile(source) as z:
        return visible_records(parse(z.read('word/document.xml')))


def _note_prefix(text):
    match=re.match(r'^\s*(?P<marker>\[\d+\]|\d+|[①-⑳]|[*†‡])(?:[.．、:：)]|\s)+',text)
    if not match:
        raise ContentIntegrityError('content-footnote-prefix-ambiguous','脚注标记与正文无法安全分离。')
    return match


def _marker_key(text):
    text=text.strip('[]')
    if len(text)==1 and '①'<=text<='⑳':return str(ord(text)-ord('①')+1)
    return text


def validate_movements(source, movements):
    records=_paragraph_records(source);seen=set()
    for move in movements:
        if set(move)!={'reference_paragraph','note_paragraph','note_before','note_body','note_id','start','end','marker'}:
            raise ContentIntegrityError('content-footnote-plan-invalid','脚注保护计划字段不完整或含未知字段。')
        ri,ni=move['reference_paragraph'],move['note_paragraph']
        if type(ri)!=int or type(ni)!=int or not 1<=ri<=len(records) or not 1<=ni<=len(records) or ri==ni:
            raise ContentIntegrityError('content-footnote-plan-invalid','脚注定位无效。')
        ref,note=records[ri-1],records[ni-1]
        note_text=''.join(note['text']);ref_text=''.join(ref['text']);prefix=_note_prefix(note_text)
        if note['node'].getparent().tag!=Q('body') or note_text!=move['note_before'] or canonical_text(note_text[prefix.end():])!=canonical_text(move['note_body']):
            raise ContentIntegrityError('content-footnote-plan-invalid','脚注正文与原文不一致。')
        start,end=move['start'],move['end']
        if type(start)!=int or type(end)!=int or not 0<=start<end<=len(ref_text) or ref_text[start:end]!=move['marker'] or _marker_key(move['marker'])!=_marker_key(prefix['marker']):
            raise ContentIntegrityError('content-footnote-plan-invalid','只允许替换与原脚注正文匹配的标记。')
        if not re.fullmatch(r'\d+',str(move['note_id'])) or (ri,start,end) in seen or ('note',ni) in seen:
            raise ContentIntegrityError('content-footnote-plan-invalid','脚注标记重复或 ID 无效。')
        seen.update({(ri,start,end),('note',ni)})
    return movements


def register_footnote_changes(source, candidate, receipt, plan):
    """Turn exact native-note conversion receipts into bounded content movements.

    A receipt cannot authorize new note prose: it is compared to the original
    definition, reference marker and the created native note, independently.
    """
    moves=[];records=_paragraph_records(source)
    with zipfile.ZipFile(candidate) as z:
        root=parse(z.read('word/footnotes.xml')) if 'word/footnotes.xml' in z.namelist() else None
    for change in receipt.get('footnote_changes',[]):
        if 'reference_paragraph' not in change or 'note_paragraph' not in change:
            continue
        ri,ni=change['reference_paragraph'],change['note_paragraph']
        if type(ri)!=int or type(ni)!=int or not 1<=ri<=len(records) or not 1<=ni<=len(records):
            raise ContentIntegrityError('content-footnote-location','脚注转换返回的原始定位无效。')
        ref,definition=records[ri-1],records[ni-1]
        original=''.join(definition['text']);prefix=_note_prefix(original)
        if original!=change.get('before'):
            raise ContentIntegrityError('content-footnote-receipt-stale','脚注修改记录不属于原文。')
        possible=[];offset=0
        for r in ref['node'].iter(Q('r')):
            text=''.join(t.text or '' for t in r.iter(Q('t')))
            rp=r.find(Q('rPr'));va=rp.find(Q('vertAlign')) if rp is not None else None
            cs=rp.find(Q('rStyle')) if rp is not None else None
            styled=(va is not None and va.get(Q('val'))=='superscript') or (cs is not None and re.search(r'footnote.*reference|脚注引用',cs.get(Q('val'),''),re.I))
            if styled and _marker_key(text)==_marker_key(prefix['marker']):possible.append((offset,offset+len(text),text))
            offset+=len(text)
        if len(possible)!=1:
            raise ContentIntegrityError('content-footnote-marker-ambiguous','脚注引用标记不能唯一定位，不能豁免业务文字变化。')
        start,end,marker=possible[0];ident=str(change['note_id'])
        note=next((n for n in root if n.tag==Q('footnote') and n.get(Q('id'))==ident),None) if root is not None else None
        body=original[prefix.end():]
        actual=' '.join(''.join(r['text']) for r in visible_records(note)) if note is not None else None
        if actual is None or canonical_text(body)!=canonical_text(actual):
            raise ContentIntegrityError('content-footnote-body-changed','移动后的脚注正文与原文不一致。')
        moves.append({'reference_paragraph':ri,'note_paragraph':ni,'note_before':original,'note_body':body,
                      'note_id':ident,'start':start,'end':end,'marker':marker})
    validate_movements(source,moves)
    updated=dict(plan);updated['allowed_footnote_moves']=moves
    return updated

def compare_sequence(before, after, *, scope):
    if before == after:
        return []
    issues = []
    sm = difflib.SequenceMatcher(a=before, b=after, autojunk=False)
    for tag, a0, a1, b0, b1 in sm.get_opcodes():
        if tag == 'equal':
            continue
        issues.append({'severity': 'error', 'code': 'unapproved-business-content-change',
                       'scope': scope, 'operation': tag, 'source_range': [a0 + 1, a1],
                       'candidate_range': [b0 + 1, b1], 'before': before[a0:a1][:5],
                       'after': after[b0:b1][:5],
                       'message': '变化不属于已识别的排版修复；必须核对或撤销，不能按格式合格放行。'})
    return issues


def audit(source, candidate, plan=None):
    source, candidate = Path(source), Path(candidate)
    data = read_json(plan) if plan is not None else make_plan(source)
    if data.get('version') != 1 or data.get('source_sha256') != digest(source):
        raise ContentIntegrityError('content-plan-stale', '内容保护基准不属于本次原始文件。')
    moves=validate_movements(source,data.get('allowed_footnote_moves',[]))
    before, after = project(source,moves), project(candidate)
    issues = compare_sequence(before['body'], after['body'], scope='正文与表格业务文字')
    issues.extend(compare_sequence(before['protected_literals'],after['protected_literals'],scope='URL、路径与代码字面量'))
    for part in sorted(set(before['other_stories']) | set(after['other_stories'])):
        issues.extend(compare_sequence(before['other_stories'].get(part, []), after['other_stories'].get(part, []), scope=part))
    issues.extend(compare_sequence(before['tracked_deletions'], after['tracked_deletions'], scope='修订删除记录'))
    used_insertions = []
    for kind in ('figure', 'table'):
        old = before['captions'][kind]
        new = after['captions'][kind]
        allowed = Counter(canonical_text(r['title']) for r in data.get('allowed_caption_insertions', []) if r.get('kind') == kind)
        # Existing captions must stay in order. Only explicitly planned additions may be removed.
        i, filtered = 0, []
        for title in new:
            if i < len(old) and title == old[i]:
                filtered.append(title)
                i += 1
            elif allowed[title] > 0:
                allowed[title] -= 1
                used_insertions.append({'kind': kind, 'title': title})
            else:
                filtered.append(title)
        issues.extend(compare_sequence(old, filtered, scope=kind + '题注业务题名'))
    return {'status': 'failed' if issues else 'passed', 'source_sha256': digest(source),
            'candidate_sha256': digest(candidate), 'issues': issues,
            'allowed_caption_insertions_used': used_insertions,'allowed_footnote_moves_used':len(moves),
            'checked_body_paragraphs': len(after['body']),
            'note': '直接对照原文；视觉和图内语义由独立逐图验收负责。已登记的脚注搬移逐字核对，纵向合并按逻辑行展开核对；其他未识别的内容变化不会自动放行。'}


def main():
    ap = argparse.ArgumentParser(description='核对原始 Word 与最终 Word 的业务内容')
    ap.add_argument('source', type=Path)
    ap.add_argument('candidate', type=Path)
    ap.add_argument('--content-plan', type=Path)
    ap.add_argument('--json-out', type=Path)
    a = ap.parse_args()
    try:
        result = audit(a.source, a.candidate, a.content_plan)
    except (ContentIntegrityError, OSError, ValueError, KeyError, zipfile.BadZipFile) as exc:
        result = {'status': 'failed', 'issues': [exc.as_issue() if isinstance(exc, ContentIntegrityError) else
                  {'severity': 'error', 'code': 'content-audit-failed', 'message': str(exc)}]}
    payload = json.dumps(result, ensure_ascii=False, indent=2)
    if a.json_out:
        a.json_out.parent.mkdir(parents=True, exist_ok=True)
        a.json_out.write_text(payload + '\n', encoding='utf-8')
    print(payload)
    return 0 if result['status'] == 'passed' else 2


if __name__ == '__main__':
    raise SystemExit(main())
