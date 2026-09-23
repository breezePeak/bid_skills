"""Audit and convert Word-native heading, caption and footnote numbering.

Only edit identified semantic objects. Ambiguous sequences/references are reported,
not guessed. All package parts not changed by a repair retain their original bytes.
"""
from __future__ import annotations

import copy
import hashlib
import json
import re
import shutil
import tempfile
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from lxml import etree as E

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
NS = {'w': W, 'r': R}
Q = lambda n: '{%s}%s' % (W, n)
FALSE = {'0', 'false', 'off'}
CN = '零〇一二三四五六七八九十百千两'
DIGITS = '零一二三四五六七八九'


class NumberingError(ValueError):
    pass


def xml(data: bytes):
    return E.fromstring(data, E.XMLParser(resolve_entities=False, no_network=True))


def dump(root) -> bytes:
    return E.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)


def val(node, name='val', default=None):
    return default if node is None else node.get(Q(name), default)


def enabled(node, attr='val'):
    return node is not None and str(val(node, attr, '1')).lower() not in FALSE


def item(tag, **attrs):
    return E.Element(Q(tag), {Q(k): str(v) for k, v in attrs.items()})


def set_child(parent, tag, **attrs):
    n = parent.find('w:' + tag, NS)
    if n is None:
        n = E.SubElement(parent, Q(tag))
    for k, v in attrs.items():
        n.set(Q(k), str(v))
    return n


def ppr(p):
    n = p.find('w:pPr', NS)
    if n is None:
        n = item('pPr'); p.insert(0, n)
    return n


def set_number(p, num_id: str, level: int):
    pp = ppr(p)
    n = pp.find('w:numPr', NS)
    if n is None:
        n = item('numPr')
        # CT_PPr: numPr precedes spacing/ind/jc and follows keep/page flags.
        later = {'suppressLineNumbers','pBdr','shd','tabs','suppressAutoHyphens',
                 'kinsoku','wordWrap','overflowPunct','topLinePunct','autoSpaceDE',
                 'autoSpaceDN','bidi','adjustRightInd','snapToGrid','spacing','ind',
                 'contextualSpacing','mirrorIndents','suppressOverlap','jc',
                 'textDirection','textAlignment','textboxTightWrap','outlineLvl',
                 'divId','cnfStyle','rPr','sectPr','pPrChange'}
        position = next((i for i, c in enumerate(pp) if E.QName(c).localname in later), len(pp))
        pp.insert(position, n)
    for tag in ('ilvl','numId'):
        for c in list(n.findall('w:' + tag, NS)): n.remove(c)
    n.insert(0, item('ilvl', val=level)); n.insert(1, item('numId', val=num_id))


def visible_nodes(p):
    """Offsets use visible w:t, tabs and breaks; field instructions are excluded."""
    offset = 0
    for n in p.iter():
        if any(a.tag in {Q('del'), Q('moveFrom')} for a in n.iterancestors()):
            continue
        value = n.text or '' if n.tag == Q('t') else ('\t' if n.tag == Q('tab') else '\n' if n.tag in {Q('br'), Q('cr')} else '')
        if value:
            yield n, offset, offset + len(value), value
            offset += len(value)


def text(p):
    return ''.join(v for _, _, _, v in visible_nodes(p))


def run(text_value='', props=None):
    r = item('r')
    if props is not None: r.append(copy.deepcopy(props))
    if text_value:
        t = E.SubElement(r, Q('t')); t.text = text_value
        if text_value[:1].isspace() or text_value[-1:].isspace():
            t.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
    return r


def field(instruction: str, cached: str, props=None):
    f = item('fldSimple', instr=' ' + instruction.strip() + ' ', dirty='true')
    f.append(run(cached, props))
    return f


def span_check(p, start, end):
    if not 0 <= start <= end <= len(text(p)):
        raise NumberingError('编号文字定位越界，请重新读取当前文件。')
    selected = [(n,a,b,v) for n,a,b,v in visible_nodes(p) if a < end and b > start]
    if start == end:
        selected = [(n,a,b,v) for n,a,b,v in visible_nodes(p) if a <= start <= b][:1]
    for n, _, _, _ in selected:
        r = n.getparent()
        if r is None or r.tag != Q('r') or r.getparent() is not p:
            raise NumberingError('编号位于域、超链接或内容控件内，需先确认替换范围。')
        if any(c.tag not in {Q('rPr'),Q('t'),Q('tab'),Q('br'),Q('cr')} for c in r):
            raise NumberingError('编号与非文本对象共用 run，需先拆分，不能整段重建。')
    return selected


def replace_span(p, start, end, replacements=()):
    """Replace a plain-text range across runs without rebuilding the paragraph.

    Bookmarks outside the replaced range and all unaffected run formatting survive.
    Refuse fields/hyperlinks rather than flattening them.
    """
    selected = span_check(p, start, end)
    if not selected:
        if start == end == len(text(p)):
            for n in replacements: p.append(n)
            return
        raise NumberingError('找不到待替换的编号文字。')
    first, a, b, value = selected[0]
    first_r = first.getparent()
    insert_at = p.index(first_r)
    if replacements and start > a:
        # Split the starting run at the exact visible boundary.
        tail = E.Element(first_r.tag, dict(first_r.attrib))
        rp = first_r.find('w:rPr', NS)
        if rp is not None: tail.append(copy.deepcopy(rp))
        children = list(first_r); position = children.index(first)
        prefix = start - a
        if first.tag != Q('t'):
            raise NumberingError('编号起点不在普通文字中。')
        suffix_node = copy.deepcopy(first); suffix_node.text = value[prefix:]
        first.text = value[:prefix]
        first.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
        suffix_node.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
        tail.append(suffix_node)
        for c in children[position+1:]:
            first_r.remove(c); tail.append(c)
        p.insert(insert_at+1, tail); insert_at += 1
    elif replacements and start == b:
        insert_at += 1
    for n, a, b, value in list(visible_nodes(p)):
        lo, hi = max(start,a)-a, min(end,b)-a
        if lo >= hi: continue
        if n.tag == Q('t'):
            n.text = value[:lo] + value[hi:]
            if (n.text or '')[:1].isspace() or (n.text or '')[-1:].isspace():
                n.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
        else:
            n.getparent().remove(n)
    for i,n in enumerate(replacements): p.insert(insert_at+i,n)


def read_fields(p):
    fields, stack, offset = [], [], 0
    for n in p.iter():
        if n.tag == Q('fldSimple'):
            cached = text(n)
            fields.append({'instr': val(n,'instr',''), 'start': offset, 'end': offset+len(cached),
                           'cached': cached, 'node': n, 'complete': True,
                           'locked': enabled(n,'fldLock') if val(n,'fldLock') is not None else False})
        elif n.tag == Q('fldChar'):
            kind = val(n,'fldCharType')
            if kind == 'begin':
                stack.append({'instr':'','start':offset,'node':n,'cached':'','separate':False,
                              'locked': enabled(n,'fldLock') if val(n,'fldLock') is not None else False})
            elif kind == 'separate' and stack: stack[-1]['separate'] = True; stack[-1]['start'] = offset
            elif kind == 'end' and stack:
                f = stack.pop(); f['end'] = offset; f['complete'] = f.pop('separate'); fields.append(f)
        elif n.tag == Q('instrText') and stack:
            stack[-1]['instr'] += n.text or ''
        elif n.tag in {Q('t'),Q('tab'),Q('br'),Q('cr')}:
            value = (n.text or '') if n.tag == Q('t') else ('\t' if n.tag == Q('tab') else '\n')
            offset += len(value)
            for f in stack:
                if f['separate']: f['cached'] += value
    for f in stack: f['end']=offset; f['complete']=False; fields.append(f)
    return fields


def command(f):
    return f['instr'].strip().split(None,1)[0].upper() if f['instr'].strip() else ''


def set_field_instruction(f,instruction):
    if f['node'].tag==Q('fldSimple'):
        f['node'].set(Q('instr'),' '+instruction.strip()+' ')
        f['node'].set(Q('dirty'),'true')
    else:
        start=f['node'];p=next(a for a in start.iterancestors() if a.tag==Q('p'))
        active=False;nodes=[]
        for n in p.iter():
            if n is start:active=True
            elif active and n.tag==Q('fldChar') and val(n,'fldCharType')=='separate':break
            elif active and n.tag==Q('instrText'):nodes.append(n)
        if not nodes:raise NumberingError('复杂域缺少指令文字。')
        nodes[0].text=' '+instruction.strip()+' '
        for n in nodes[1:]:n.text=''
        start.set(Q('dirty'),'true')
    f['instr']=instruction


def chinese_int(s):
    if s.isdecimal(): return int(s)
    digits = {c:i for i,c in enumerate(DIGITS)}; digits.update({'〇':0,'两':2})
    total, number = 0, 0
    for c in s:
        if c in digits: number = digits[c]
        elif c in {'十','百','千'}: total += (number or 1)*{'十':10,'百':100,'千':1000}[c]; number=0
        else: raise NumberingError('不支持的中文序号：'+s)
    return total+number


def number_text(n, fmt):
    if fmt == 'decimal': return str(n)
    if fmt == 'decimalZero': return str(n).zfill(2)
    if fmt in {'chineseCounting','chineseCountingThousand'}:
        if not 0 <= n < 10000: raise NumberingError('中文标题序号超出自动转换范围。')
        if n == 0: return '零'
        out=''; zero=False
        for unit, label in ((1000,'千'),(100,'百'),(10,'十'),(1,'')):
            digit,n=divmod(n,unit)
            if digit:
                if zero: out+='零'; zero=False
                out+=DIGITS[digit]+label
            elif out and n: zero=True
        return out[1:] if out.startswith('一十') else out
    raise NumberingError('现有编号格式需由 Agent 确认：'+str(fmt))


def manual_heading(s, level):
    m = re.match(r'^\s*第(?P<n>[0-9'+CN+r']+)(?P<unit>章|节|篇|部分)\s*',s)
    if m:
        n=m['n']; return {'end':m.end(),'values':[chinese_int(n)],'fmt':'decimal' if n.isdecimal() else 'chineseCounting',
                         'pattern':'第%'+str(level+1)+m['unit'], 'token':m[0].strip(), 'suffix':'space'}
    m = re.match(r'^\s*(?P<open>[（(]?)(?P<n>['+CN+r']+)(?P<close>[）)．.、])\s*',s)
    if m:
        return {'end':m.end(),'values':[chinese_int(m['n'])], 'fmt':'chineseCounting',
                'pattern':m['open']+'%'+str(level+1)+m['close'],'token':m[0].strip(),'suffix':'space'}
    m = re.match(r'^\s*(?P<open>[（(]?)(?P<n>\d+(?:[.．]\d+)*)(?P<close>[）)．.、]?)(?P<space>\s*)',s)
    if m and (m['close'] or m['space'] or '.' in m['n'] or '．' in m['n'] or m.end()==len(s)):
        nums = [int(v) for v in re.split('[.．]',m['n'])]
        if len(nums) not in {1,level+1}: return {'error':'手工编号层级与标题样式不一致。'}
        tokens = ['%'+str(i+1) for i in range(len(nums))] if len(nums)>1 else ['%'+str(level+1)]
        separator = '．' if '．' in m['n'] else '.'
        return {'end':m.end(),'values':nums,'fmt':'decimal','pattern':m['open']+separator.join(tokens)+m['close'],
                'token':m[0].strip(),'suffix':'space'}
    return None


CAPTION = re.compile(r'^\s*(?P<label>Figure|Table|图|表)\s*(?P<number>\d+(?:[.．\-－]\d+)*)(?=[^0-9.．\-－]|$)',re.I)
NOTE_PREFIX = re.compile(r'^\s*(?P<marker>\[\d+\]|［\d+］|\d+[.．、]?|[①-⑳])(?=\s|[\u4e00-\u9fffA-Za-z]|$)\s*')


def marker_key(value):
    value=value.strip().strip('[]［］.．、')
    if len(value)==1 and '①'<=value<='⑳': return str(ord(value)-ord('①')+1)
    trans=str.maketrans('⁰¹²³⁴⁵⁶⁷⁸⁹','0123456789')
    value=value.translate(trans)
    return str(int(value)) if value.isdecimal() else value


class Package:
    def __init__(self, source: Path):
        self.source=Path(source)
        with zipfile.ZipFile(source) as z:
            self.infos=z.infolist()
            if len({i.filename for i in self.infos})!=len(self.infos): raise NumberingError('DOCX 中存在重复部件。')
            self.files={i.filename:z.read(i.filename) for i in self.infos}
        self.roots={}; self.changed=set()
        self.document=self.get('word/document.xml')
        self.styles=self.get('word/styles.xml')
        self.numbering=self.get('word/numbering.xml','numbering')
        self.style_map={val(n,'styleId'):n for n in self.styles.findall('w:style',NS)}
        self.paragraphs=self.document.xpath('.//w:p[not(ancestor::w:del or ancestor::w:moveFrom)]',namespaces=NS)
        self.index={p:i for i,p in enumerate(self.paragraphs,1)}

    def get(self, part, create=None):
        if part not in self.roots:
            if part in self.files: self.roots[part]=xml(self.files[part])
            elif create: self.roots[part]=E.Element(Q(create),nsmap={'w':W})
            else: raise NumberingError('DOCX 缺少部件：'+part)
        return self.roots[part]

    def touch(self, part): self.changed.add(part)

    def chain(self,p):
        sid=val(p.find('w:pPr/w:pStyle',NS))
        result=[]; seen=set()
        while sid and sid not in seen and sid in self.style_map:
            seen.add(sid); n=self.style_map[sid]; result.append(n); sid=val(n.find('w:basedOn',NS))
        return list(reversed(result))

    def names(self,p):
        return [val(n,'styleId','') for n in self.chain(p)]+[val(n.find('w:name',NS),default='') for n in self.chain(p)]

    def level(self,p):
        if p.xpath('ancestor::w:txbxContent|ancestor::w:tc',namespaces=NS): return None
        names=self.names(p)
        if any(re.match(r'^(TOC|目录|Title$|Subtitle$|封面)',n,re.I) for n in names): return None
        for prop in [p.find('w:pPr',NS)]+[n.find('w:pPr',NS) for n in reversed(self.chain(p))]:
            n=prop.find('w:outlineLvl',NS) if prop is not None else None
            if n is not None:
                value=int(val(n)); return value if 0<=value<=8 else None
        for name in reversed(names):
            m=re.fullmatch(r'(?:Heading|标题)\s*([1-9])',name,re.I)
            if m: return int(m[1])-1
        return None

    def num(self,p):
        number=None; level=None; direct=False
        for n in [s.find('w:pPr/w:numPr',NS) for s in self.chain(p)]+[p.find('w:pPr/w:numPr',NS)]:
            if n is None: continue
            v=val(n.find('w:numId',NS)); lv=val(n.find('w:ilvl',NS))
            if v is not None: number=v
            if n is p.find('w:pPr/w:numPr',NS) and lv is not None: level=int(lv); direct=True
        if number in (None,'0'): return None
        num=self.numbering.find(f'w:num[@w:numId="{number}"]',NS)
        if num is None: return {'num_id':number,'level':level,'valid':False,'reason':'numId 对应定义不存在'}
        aid=val(num.find('w:abstractNumId',NS))
        abstract=self.numbering.find(f'w:abstractNum[@w:abstractNumId="{aid}"]',NS)
        if abstract is None: return {'num_id':number,'level':level,'valid':False,'reason':'abstractNum 定义不存在'}
        visited={aid}
        while abstract.find('w:numStyleLink',NS) is not None:
            link=val(abstract.find('w:numStyleLink',NS));style=self.style_map.get(link)
            linked_num_id=val(style.find('w:pPr/w:numPr/w:numId',NS)) if style is not None else None
            linked_num=self.numbering.find(f'w:num[@w:numId="{linked_num_id}"]',NS) if linked_num_id else None
            linked_aid=val(linked_num.find('w:abstractNumId',NS)) if linked_num is not None else None
            linked_abs=self.numbering.find(f'w:abstractNum[@w:abstractNumId="{linked_aid}"]',NS) if linked_aid else None
            if linked_abs is None or linked_aid in visited:
                return {'num_id':number,'level':level,'valid':False,'reason':'编号样式链接缺失或循环'}
            visited.add(linked_aid);abstract=linked_abs
        if not direct:
            ids={val(s,'styleId') for s in self.chain(p)}
            candidates=[int(val(l,'ilvl')) for l in abstract.findall('w:lvl',NS) if val(l.find('w:pStyle',NS)) in ids]
            level=candidates[-1] if candidates else (self.level(p) or 0)
        lvl=num.find(f'w:lvlOverride[@w:ilvl="{level}"]/w:lvl',NS)
        if lvl is None: lvl=abstract.find(f'w:lvl[@w:ilvl="{level}"]',NS)
        pattern=val(lvl.find('w:lvlText',NS),default='') if lvl is not None else ''
        fmt=val(lvl.find('w:numFmt',NS)) if lvl is not None else None
        valid=lvl is not None and fmt not in (None,'none','bullet') and re.search(r'%[1-9]',pattern) is not None
        return {'num_id':number,'level':level,'valid':bool(valid),'num':num,'abstract':abstract,'lvl':lvl,
                'reason':'' if valid else '编号级别、格式或占位符无效'}

    def is_caption(self,p):
        if any(re.search(r'caption|题注|图题|表题',n,re.I) for n in self.names(p)): return True
        if not CAPTION.match(text(p)): return False
        for sibling in (p.getprevious(),p.getnext()):
            if sibling is not None and (sibling.tag==Q('tbl') or sibling.xpath('.//w:drawing|.//w:pict',namespaces=NS)):
                return True
        return False

    def is_note_text(self,p):
        return any(re.search(r'footnote\s*text|脚注文本',n,re.I) for n in self.names(p))

    def ensure_part(self, part, kind):
        relpart='word/_rels/document.xml.rels'
        if relpart in self.files: rels=self.get(relpart)
        else:
            rels=E.Element('{%s}Relationships'%REL,nsmap={None:REL}); self.roots[relpart]=rels
        matches=[n for n in rels if n.get('Type')==R+'/'+kind]
        expected_target=part.removeprefix('word/')
        if matches and (len(matches)!=1 or matches[0].get('Target') not in {expected_target,'/'+part} or matches[0].get('TargetMode')=='External'):
            raise NumberingError('现有 '+kind+' 部件关系不是标准内部路径，不能覆盖。')
        if not matches:
            ids={n.get('Id') for n in rels}; i=1
            while 'rId'+str(i) in ids: i+=1
            E.SubElement(rels,'{%s}Relationship'%REL,Id='rId'+str(i),Type=R+'/'+kind,Target=expected_target)
            self.touch(relpart)
        if '[Content_Types].xml' not in self.files: raise NumberingError('DOCX 缺少 [Content_Types].xml。')
        ct=self.get('[Content_Types].xml'); content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.'+kind+'+xml'
        node=next((n for n in ct if n.get('PartName')=='/'+part),None)
        if node is None:
            E.SubElement(ct,'{%s}Override'%CT,PartName='/'+part,ContentType=content_type);self.touch('[Content_Types].xml')
        elif node.get('ContentType')!=content_type:
            raise NumberingError(kind+' 部件 ContentType 不匹配。')

    def part_registered(self,part,kind):
        try:
            rels=self.get('word/_rels/document.xml.rels')
            ok=any(n.get('Type')==R+'/'+kind and n.get('Target') in {part.removeprefix('word/'),'/'+part} and n.get('TargetMode')!='External' for n in rels)
            ct=self.get('[Content_Types].xml')
            return ok and any(n.get('PartName')=='/'+part and n.get('ContentType')=='application/vnd.openxmlformats-officedocument.wordprocessingml.'+kind+'+xml' for n in ct)
        except NumberingError: return False

    def update_fields(self):
        settings=self.get('word/settings.xml','settings')
        existing=settings.find('w:updateFields',NS)
        if existing is None:
            existing=item('updateFields',val='true')
            later={'hdrShapeDefaults','footnotePr','endnotePr','compat','docVars','rsids','mathPr',
                   'attachedSchema','themeFontLang','clrSchemeMapping','doNotIncludeSubdocsInStats',
                   'doNotAutoCompressPictures','forceUpgrade','captions','readModeInkLockDown',
                   'smartTagType','shapeDefaults','doNotEmbedSmartTags','decimalSymbol','listSeparator'}
            at=next((i for i,n in enumerate(settings) if E.QName(n).localname in later),len(settings))
            settings.insert(at,existing)
        else:existing.set(Q('val'),'true')
        self.touch('word/settings.xml')
        self.ensure_part('word/settings.xml','settings')

    def save(self,out):
        out=Path(out)
        if out.resolve()==self.source.resolve(): raise NumberingError('输出文件不得覆盖输入文档。')
        out.parent.mkdir(parents=True,exist_ok=True)
        if not self.changed: shutil.copy2(self.source,out);return
        files=dict(self.files)
        for name in self.changed: files[name]=dump(self.roots[name])
        with tempfile.NamedTemporaryFile(dir=out.parent,suffix='.docx',delete=False) as tmp: tmp_path=Path(tmp.name)
        try:
            with zipfile.ZipFile(tmp_path,'w',zipfile.ZIP_DEFLATED) as z:
                for name,data in files.items(): z.writestr(name,data)
            tmp_path.replace(out)
        finally:
            if tmp_path.exists():tmp_path.unlink()


def heading_snapshot(package):
    return [(i,p,package.level(p),package.num(p)) for i,p in enumerate(package.paragraphs,1) if package.level(p) is not None and (text(p).strip() or package.num(p))]


def note_candidates(pkg):
    definitions=[];refs=defaultdict(list)
    for p in pkg.paragraphs:
        if pkg.is_note_text(p):
            m=NOTE_PREFIX.match(text(p))
            if m: definitions.append((p,m,marker_key(m['marker'])))
            continue
        for r in p.findall('w:r',NS):
            value=text(r).strip()
            if not value or r.find('w:footnoteReference',NS) is not None: continue
            style=val(r.find('w:rPr/w:rStyle',NS),default='')
            superscript=val(r.find('w:rPr/w:vertAlign',NS))=='superscript'
            if re.fullmatch(r'(?:\[\d+\]|［\d+］|\d+|[①-⑳]|[⁰¹²³⁴⁵⁶⁷⁸⁹]+)',value) and (superscript or re.search(r'footnote\s*reference|脚注引用',style,re.I)):
                refs[marker_key(value)].append((p,r,value))
    return definitions,refs


def issue(code,message,paragraph=None,severity='error',**extra):
    return {'severity':severity,'code':code,'message':message,**({'paragraph':paragraph} if paragraph else {}),**extra}


def audit_package(pkg):
    issues=[];counts={'headings':0,'captions':0,'footnotes':0,'manual_footnotes':0}
    for i,p,level,num in heading_snapshot(pkg):
        counts['headings']+=1
        if not num or not num['valid']:
            issues.append(issue('heading-not-automatic','标题未使用有效的 Word 多级自动编号。',i,level=level+1,text=text(p)[:120]))
        elif num['level']!=level:
            issues.append(issue('heading-number-level-mismatch','自动编号级别与标题层级不一致。',i))
        if num and num['valid'] and not pkg.part_registered('word/numbering.xml','numbering'):
            issues.append(issue('numbering-part-unregistered','自动编号部件未正确注册。',i))
        if num and num['valid'] and manual_heading(text(p),level):
            issues.append(issue('heading-manual-plus-automatic','标题同时存在自动编号和手工编号文字，需核对是否重复。',i,severity='review'))
    seqs=defaultdict(list)
    for i,p in enumerate(pkg.paragraphs,1):
        if not pkg.is_caption(p):continue
        counts['captions']+=1
        fields=read_fields(p); seq=[f for f in fields if command(f)=='SEQ']
        if len(seq)!=1:
            issues.append(issue('caption-not-automatic','题注必须有一个有效的 SEQ 自动序号域；普通数字不算自动编号。',i,text=text(p)[:120]));continue
        f=seq[0]
        m=re.match(r'\s*SEQ\s+(?:"([^"]+)"|([^\s\\]+))',f['instr'],re.I)
        if not f['complete'] or not m or f['locked'] or re.search(r'\\[ch]\b',f['instr'],re.I):
            issues.append(issue('caption-sequence-invalid','题注序号域未闭合、被锁定、缺少序列名或仅引用/隐藏编号。',i));continue
        mcap=CAPTION.match(text(p))
        if not mcap or not (mcap.start('number')<=f['start']<f['end']<=mcap.end('number')):
            issues.append(issue('caption-sequence-not-label','SEQ 域不在题注编号位置，或域结果没有有效序号。',i));continue
        raw=mcap['number'];parts=re.split(r'[.．\-－]',raw)
        if len(parts)>1:
            style_refs=[x for x in fields if command(x)=='STYLEREF' and x['complete'] and not x['locked'] and x['start']==mcap.start('number') and x['end']<f['start'] and re.search(r'\\[snrw]\b',x['instr'],re.I)]
            if not style_refs or not re.search(r'\\s\s+[1-9]\b',f['instr'],re.I):
                issues.append(issue('caption-chapter-prefix-static','分章题注的章节号也必须自动引用标题，序号应按对应标题级别重启。',i))
        if mcap.end('number')!=f['end']:
            issues.append(issue('caption-static-number-tail','题注自动域之后仍有手工编号片段。',i))
        seqs[(m[1] or m[2]).casefold()].append((i,f))
    for name,fields in seqs.items():
        if len(fields)>1 and all(re.search(r'\\r\s+\d+',f['instr'],re.I) for _,f in fields):
            issues.append(issue('caption-every-number-reset','同一题注序列逐条固定重启，不能随增删正常连续编号。',fields[0][0],sequence=name))
    body_refs=pkg.document.findall('.//w:footnoteReference',NS)
    if body_refs or 'word/footnotes.xml' in pkg.files or 'word/footnotes.xml' in pkg.changed:
        if 'word/footnotes.xml' not in pkg.files and 'word/footnotes.xml' not in pkg.changed:
            issues.append(issue('footnote-part-missing','正文存在脚注引用，但脚注内容部件缺失。'))
        else:
            notes=pkg.get('word/footnotes.xml');regular=[n for n in notes.findall('w:footnote',NS) if val(n,'type','normal')=='normal']
            ids=[val(n,'id') for n in regular]; refs=[val(n,'id') for n in body_refs]
            counts['footnotes']=len(regular)
            if not pkg.part_registered('word/footnotes.xml','footnotes'):
                issues.append(issue('footnotes-part-unregistered','脚注部件关系或内容类型缺失。'))
            for ident,count in Counter(ids).items():
                if count!=1:issues.append(issue('footnote-id-duplicate','脚注 ID 重复。',note_id=ident))
            for ident,count in Counter(refs).items():
                if ident not in ids:issues.append(issue('footnote-reference-broken','脚注引用找不到对应内容。',note_id=ident))
                if count>1:issues.append(issue('footnote-reference-duplicate','同一脚注被多个原生脚注引用重复使用，需核对交叉引用。',note_id=ident))
            for n in regular:
                if val(n,'id') not in refs:issues.append(issue('footnote-orphan','脚注内容没有正文引用，不能猜测配对。',note_id=val(n,'id')))
                if len(n.findall('.//w:footnoteRef',NS))!=1:issues.append(issue('footnote-number-not-automatic','脚注正文缺少唯一的原生自动编号标记。',note_id=val(n,'id')))
            for r in body_refs:
                if val(r,'customMarkFollows') is not None and enabled(r,'customMarkFollows'):
                    issues.append(issue('footnote-custom-mark','脚注使用自定义手工标记，尚非自动序号。',note_id=val(r,'id')))
    definitions,refs=note_candidates(pkg)
    definition_keys={key for _,_,key in definitions}
    for p,m,key in definitions:
        counts['manual_footnotes']+=1
        issues.append(issue('footnote-manual-pair' if len(refs[key])==1 else 'footnote-manual-ambiguous',
                            '脚注是普通文字/上标，应转换为原生脚注引用和内容。' if len(refs[key])==1 else '手工脚注引用不能唯一匹配，需 Agent 确认位置。',
                            pkg.index[p],severity='error' if len(refs[key])==1 else 'review',marker=m['marker']))
    for key,items in refs.items():
        if key not in definition_keys:
            for p,r,value in items:
                style=val(r.find('w:rPr/w:rStyle',NS),default='')
                if re.search(r'footnote\s*reference|脚注引用',style,re.I):
                    issues.append(issue('footnote-manual-reference-unmatched','疑似手工脚注引用没有对应脚注内容，不能按指数或普通数字自动删除。',pkg.index[p],severity='review',marker=value))
    return {'status':'passed' if not issues else 'failed','counts':counts,'issues':issues,
            'visual_review_required':True,'note':'仅程序结构检查；仍需更新域并渲染复查实际编号和分页。'}


def build_heading_numbering(pkg, records):
    patterns={};starts={}; style_ids={}
    for i,p,level,_ in records:
        manual=manual_heading(text(p),level)
        if manual and 'error' in manual: raise NumberingError(manual['error']+' 段落 '+str(i))
        if manual:
            shape=(manual['fmt'],manual['pattern'])
            if level in patterns and patterns[level]!=shape:
                raise NumberingError('同级标题编号格式不一致，请先确认。段落 '+str(i))
            patterns[level]=shape;starts.setdefault(level,manual['values'][-1])
        sid=val(p.find('w:pPr/w:pStyle',NS))
        if sid: style_ids.setdefault(level,sid)
    ids=[int(val(n,'abstractNumId')) for n in pkg.numbering.findall('w:abstractNum',NS)]
    aid=max(ids,default=-1)+1
    absn=item('abstractNum',abstractNumId=aid);absn.append(item('multiLevelType',val='multilevel'))
    for level in range(9):
        fmt,pattern=patterns.get(level,('decimal','.'.join('%'+str(i+1) for i in range(level+1))))
        lvl=item('lvl',ilvl=level)
        lvl.append(item('start',val=starts.get(level,1)))
        lvl.append(item('numFmt',val=fmt))
        if level in style_ids:lvl.append(item('pStyle',val=style_ids[level]))
        lvl.append(item('suff',val='space'));lvl.append(item('lvlText',val=pattern));lvl.append(item('lvlJc',val='left'))
        # Do not introduce list indents: paragraph/template remains the layout baseline.
        absn.append(lvl)
    first_num=next((i for i,n in enumerate(pkg.numbering) if n.tag==Q('num')),len(pkg.numbering))
    pkg.numbering.insert(first_num,absn)
    ids=[int(val(n,'numId')) for n in pkg.numbering.findall('w:num',NS)]
    nid=str(max(ids,default=0)+1);num=item('num',numId=nid);num.append(item('abstractNumId',val=aid));pkg.numbering.append(num)
    pkg.touch('word/numbering.xml');pkg.ensure_part('word/numbering.xml','numbering')
    return nid


def heading_values(pkg, records):
    states={};result={}
    for i,p,level,num in records:
        num=pkg.num(p)
        if not num or not num['valid'] or num['level']!=level:raise NumberingError('标题编号结构尚未有效。段落 '+str(i))
        nid=num['num_id'];state=states.setdefault(nid,[None]*9)
        lvl=num['lvl'];start=int(val(lvl.find('w:start',NS),default='1'))
        override=num['num'].find(f'w:lvlOverride[@w:ilvl="{level}"]/w:startOverride',NS)
        if state[level] is None and override is not None:start=int(val(override))
        state[level]=start if state[level] is None else state[level]+1
        # Word's default lower-level restart is triggered by the preceding level.
        for deeper in range(level+1,9):
            dl=num['abstract'].find(f'w:lvl[@w:ilvl="{deeper}"]',NS)
            restart=val(dl.find('w:lvlRestart',NS),default=str(deeper)) if dl is not None else str(deeper)
            if int(restart)!=0 and level<int(restart):state[deeper]=None
        pattern=val(lvl.find('w:lvlText',NS),default='')
        def replace(m):
            lv=int(m[1])-1
            if state[lv] is None: raise NumberingError('标题跳级或缺少上级编号，不能补造。段落 '+str(i))
            definition=num['abstract'].find(f'w:lvl[@w:ilvl="{lv}"]',NS)
            fmt=val(definition.find('w:numFmt',NS),default='decimal') if definition is not None else 'decimal'
            return number_text(state[lv],fmt)
        label=re.sub(r'%([1-9])',replace,pattern)
        result[p]={'label':label,'values':list(state),'num_id':nid,'level':level,
                   'style_name':next((val(s.find('w:name',NS)) for s in reversed(pkg.chain(p)) if val(s.find('w:name',NS))),None),
                   'format':val(lvl.find('w:numFmt',NS),default='decimal')}
    return result


def repair_headings(pkg,changes):
    records=heading_snapshot(pkg)
    if not records:return
    needs=[r for r in records if not r[3] or not r[3]['valid'] or manual_heading(text(r[1]),r[2])]
    if not needs:return
    if any(read_fields(p) for _,p,_,_ in needs):raise NumberingError('待转换标题含有域，需先确认原编号范围。')
    valid_ids={r[3]['num_id'] for r in records if r[3] and r[3]['valid']}
    if len(valid_ids)>1:raise NumberingError('标题使用多个独立编号序列，需 Agent 确认手工标题应加入哪个序列。')
    if valid_ids:
        nid=next(iter(valid_ids))
    else:
        nid=build_heading_numbering(pkg,records)
    for _,p,level,_ in needs:set_number(p,nid,level)
    values=heading_values(pkg,records)
    # Preflight all prefixes; never use per-paragraph restart to disguise a bad sequence.
    for i,p,level,_ in needs:
        prefix=manual_heading(text(p),level)
        if prefix and ('error' in prefix or prefix['token']!=values[p]['label']):
            raise NumberingError('手工标题序号与自动序列不一致，需先核对断号、重复号或引用。段落 '+str(i))
        if prefix:
            span_check(p,0,prefix['end'])
            names=[val(b,'name') for b in p.findall('w:bookmarkStart',NS)]
            for paragraph in pkg.paragraphs:
                for f in read_fields(paragraph):
                    if command(f)=='REF' and any(re.search(r'\b'+re.escape(name)+r'\b',f['instr']) for name in names if name) and not re.search(r'\\[nrw]\b',f['instr']):
                        raise NumberingError('手工标题存在按文字引用的 REF 域，需先确认引用格式，不能造成引用丢号。')
    for i,p,level,_ in needs:
        before=text(p);prefix=manual_heading(before,level)
        if prefix:replace_span(p,0,prefix['end'])
        changes.append({'kind':'heading','paragraph':i,'before':before,'number':values[p]['label'],'action':'改为 Word 多级自动编号'})
    pkg.touch('word/document.xml')


def ensure_chapter_resets(pkg, sequence_ids, changes):
    r"""Anchor each generated chapter sequence to its chapter, not to a caption.

    The normal SEQ caption still has \s 1. A hidden native reset at the heading
    also supports readers that ignore that switch, and moves with the chapter.
    Never add a per-caption fixed reset (inserting a new caption must renumber).
    """
    if not sequence_ids:return
    for _,p,level,num in heading_snapshot(pkg):
        if level != 0 or not num or not num['valid']:continue
        existing=read_fields(p)
        for ident in sorted(sequence_ids):
            reset='SEQ '+ident+' \\r 0 \\h'
            if any(f['instr'].strip().casefold()==reset.casefold() for f in existing):continue
            props=item('rPr');props.append(item('vanish'))
            f=field(reset,'',props)
            # Keep an empty result run with explicit hidden formatting: readers
            # that ignore \h must not print the reset's numeric result.
            f[0].append(item('t'))
            p.append(f)
            changes.append({'kind':'caption','paragraph':pkg.index[p],
                            'sequence':ident,'action':'在章节标题中建立隐藏的自动序列重启域'})
            pkg.touch('word/document.xml')


def repair_captions(pkg,changes):
    records=heading_snapshot(pkg)
    try: values=heading_values(pkg,records) if records else {}
    except NumberingError:values={}
    chapter=None; chapter_key=None; data=[]; seq_by_label={}; existing=defaultdict(list)
    for p in pkg.paragraphs:
        if pkg.level(p)==0:
            chapter=values.get(p);chapter_key=p
        if not pkg.is_caption(p):continue
        s=text(p);m=CAPTION.match(s);fields=read_fields(p);seq=[f for f in fields if command(f)=='SEQ']
        if not m:
            raise NumberingError('题注缺少可确定的编号位置，需先确认标签和编号格式。段落 '+str(pkg.index[p]))
        label=m['label'].casefold();parts=re.split(r'[.．\-－]',m['number']);separator=re.search(r'[.．\-－]',m['number'])
        if len(parts)>2:raise NumberingError('多级章节题注需 Agent 确认章节来源，不能猜测。')
        if seq:
            ident=re.match(r'\s*SEQ\s+(?:"([^"]+)"|([^\s\\]+))',seq[0]['instr'],re.I)
            if len(seq)!=1 or not ident or not seq[0]['complete']:raise NumberingError('现有题注域结构异常，需先修复。')
            seq_id=ident[1] or ident[2]
            if label in seq_by_label and seq_by_label[label]!=seq_id:raise NumberingError('相同题注标签使用不同 SEQ 序列，需先确认。')
            seq_by_label[label]=seq_id;existing[label].append(seq[0])
        data.append((p,m,parts,separator[0] if separator else '',chapter,chapter_key,seq))
    reset_ids=set()
    # Unchanged, valid automatic captions need no sequence reconstruction.
    for p,m,parts,separator,ch,ch_key,seq in data:
        if not seq:continue
        f=seq[0]
        if f['locked']:
            f['node'].attrib.pop(Q('fldLock'),None)
            changes.append({'kind':'caption','paragraph':pkg.index[p],'action':'解除题注序号域锁定'})
            pkg.touch('word/document.xml')
        if len(parts)==2 and not any(command(x)=='STYLEREF' for x in read_fields(p)):
            if ch is None or ch['values'][0]!=int(parts[0]) or ch['format'] not in {'decimal','decimalZero'}:
                raise NumberingError('题注章节号与标题编号形式不一致，需先确认。')
            start=m.start('number');end=start+len(parts[0])
            props=span_check(p,start,end)[0][0].getparent().find('w:rPr',NS)
            style=ch.get('style_name') or 'Heading 1'
            replace_span(p,start,end,[field('STYLEREF "'+style+'" \\n \\t',parts[0],props)])
            if not re.search(r'\\s\s+1\b',f['instr'],re.I):
                if re.search(r'\\s\s+',f['instr'],re.I):raise NumberingError('题注重启层级与章节来源不一致。')
                set_field_instruction(f,f['instr'].strip()+' \\s 1')
            reset_ids.add(seq_by_label[m['label'].casefold()])
            changes.append({'kind':'caption','paragraph':pkg.index[p],'action':'手工章节号改为标题引用域'})
            pkg.touch('word/document.xml')
    if data and all(seq for *_,seq in data):
        ensure_chapter_resets(pkg,reset_ids,changes)
        return
    counters={};modes={};plans=[]
    for p,m,parts,separator,ch,ch_key,seq in data:
        label=m['label'].casefold();mode=len(parts)
        if label in modes and modes[label]!=mode:raise NumberingError('同类题注混用连续编号与分章编号，需先确认。')
        modes[label]=mode
        if mode==2:
            if ch is None or ch['values'][0]!=int(parts[0]) or ch['format'] not in {'decimal','decimalZero'}:raise NumberingError('题注章节号与上级标题不匹配，不能写死章节号。')
        key=(label,ch_key if mode==2 else None)
        wanted=int(parts[-1]); expected=counters.get(key,0)+1
        if key not in counters and mode==1: expected=wanted  # one initial restart is legitimate
        if wanted!=expected:raise NumberingError('题注存在断号、重复号或顺序冲突，请先核对引用。段落 '+str(pkg.index[p]))
        counters[key]=expected
        if seq:continue
        if read_fields(p):raise NumberingError('手工题注范围混有其他域，需先确认，不能重建整段。')
        if any(re.search(r'\\[crh]\b',f['instr'],re.I) for f in existing[label]):
            raise NumberingError('已有题注序列含重启或引用开关，需确认后再加入手工题注。')
        span_check(p,m.start('number'),m.end('number'))
        seq_id=seq_by_label.get(label,'Figure' if label in {'图','figure'} else 'Table')
        instruction='SEQ '+seq_id+' \\* ARABIC'
        if mode==2:
            instruction+=' \\s 1'
            reset_ids.add(seq_id)
        elif expected!=1 and not any(x[1]['label'].casefold()==label for x in plans) and not existing[label]:instruction+=' \\r '+str(expected)
        plans.append((p,m,instruction,parts,separator,ch))
    for p,m,instruction,parts,separator,ch in plans:
        before=text(p);nodes=span_check(p,m.start('number'),m.end('number'))
        props=nodes[0][0].getparent().find('w:rPr',NS) if nodes else None
        replacements=[]
        if len(parts)==2:
            replacements.extend([field('STYLEREF "'+(ch.get('style_name') or 'Heading 1')+'" \\n \\t',parts[0],props),run(separator,props)])
        replacements.append(field(instruction,parts[-1],props))
        replace_span(p,m.start('number'),m.end('number'),replacements)
        changes.append({'kind':'caption','paragraph':pkg.index[p],'before':before,'action':'手工序号改为 SEQ 域；分章编号同步引用标题'})
    if plans:pkg.touch('word/document.xml')
    ensure_chapter_resets(pkg,reset_ids,changes)


def footnote_ref_run(ident,inside=False):
    r=run();rp=item('rPr');rp.append(item('rStyle',val='FootnoteReference'));rp.append(item('vertAlign',val='superscript'));r.append(rp)
    r.append(item('footnoteRef') if inside else item('footnoteReference',id=ident));return r


def ensure_footnotes(pkg):
    notes=pkg.get('word/footnotes.xml','footnotes')
    for ident,kind,tag in ((-1,'separator','separator'),(0,'continuationSeparator','continuationSeparator')):
        if not any(val(n,'type')==kind for n in notes.findall('w:footnote',NS)):
            if any(val(n,'id')==str(ident) for n in notes.findall('w:footnote',NS)):
                raise NumberingError('脚注分隔符 ID 被普通脚注占用，需先修复。')
            note=item('footnote',id=ident,type=kind);p=item('p');r=item('r');r.append(item(tag));p.append(r);note.append(p);notes.insert(0,note)
    pkg.ensure_part('word/footnotes.xml','footnotes');pkg.touch('word/footnotes.xml')
    return notes


def move_note_relationships(pkg,p):
    attrs=[(n,k,v) for n in p.iter() for k,v in n.attrib.items() if k in {'{'+R+'}id','{'+R+'}embed','{'+R+'}link'}]
    if not attrs:return
    src=pkg.get('word/_rels/document.xml.rels');name='word/_rels/footnotes.xml.rels'
    dst=pkg.get(name) if name in pkg.files or name in pkg.roots else E.Element('{%s}Relationships'%REL,nsmap={None:REL})
    pkg.roots[name]=dst;mapping={}
    for n,key,value in attrs:
        if value not in mapping:
            original=next((r for r in src if r.get('Id')==value),None)
            if original is None:raise NumberingError('手工脚注中的超链接或图片关系缺失。')
            used={r.get('Id') for r in dst};i=1
            while 'rId'+str(i) in used:i+=1
            fresh=copy.deepcopy(original);fresh.set('Id','rId'+str(i));dst.append(fresh);mapping[value]=fresh.get('Id')
        n.set(key,mapping[value])
    pkg.touch(name)


def repair_footnotes(pkg,changes,explicit=()):
    definitions,refs=note_candidates(pkg)
    plans=[];taken=set()
    for entry in explicit:
        rp=pkg.paragraphs[entry['reference_paragraph']-1];np=pkg.paragraphs[entry['note_paragraph']-1]
        start,end=entry['start'],entry['end'];marker=text(rp)[start:end]
        m=NOTE_PREFIX.match(text(np))
        if not m or marker_key(m['marker'])!=marker_key(marker):raise NumberingError('脚注正文序号与引用标记不一致。')
        plans.append((rp,start,end,np,m));taken.add(np)
    for np,m,key in definitions:
        if np in taken:continue
        if sum(1 for _,_,other in definitions if other==key)!=1 or len(refs[key])!=1:
            raise NumberingError('手工脚注不能唯一匹配，需提供经确认的 numbering-plan。段落 '+str(pkg.index[np]))
        rp,r,value=refs[key][0]
        start=next(a for n,a,b,v in visible_nodes(rp) if n.getparent() is r)
        plans.append((rp,start,start+len(text(r)),np,m));taken.add(np)
    # Preflight before moving any paragraphs.
    seen=set();seen_notes=set()
    for rp,start,end,np,m in plans:
        key=(rp,start,end)
        if key in seen or np in seen_notes or rp is np or np.getparent().tag!=Q('body'):
            raise NumberingError('脚注引用重复或脚注内容位置不安全。')
        seen.add(key);seen_notes.add(np);span_check(rp,start,end);span_check(np,0,m.end())
        if np.xpath('.//w:sectPr|.//w:footnoteReference|.//w:bookmarkStart|.//w:bookmarkEnd|.//w:drawing|.//w:pict',namespaces=NS):
            raise NumberingError('手工脚注含分节、嵌套引用、书签或图片，需先确认移动范围。')
    native=pkg.document.findall('.//w:footnoteReference',NS)
    if not plans and not native:return
    if native and 'word/footnotes.xml' not in pkg.files:raise NumberingError('脚注内容部件缺失，不能补造脚注内容。')
    notes=ensure_footnotes(pkg)
    ids=[val(n,'id') for n in notes.findall('w:footnote',NS)]
    if len(set(ids))!=len(ids):raise NumberingError('脚注 ID 重复，不能自动猜测引用关系。')
    # Custom note marks need semantic confirmation; native ordinary marks can be restored.
    for ref in native:
        ident=val(ref,'id');note=next((n for n in notes if val(n,'id')==ident and val(n,'type','normal')=='normal'),None)
        if note is None:raise NumberingError('脚注引用内容缺失：'+str(ident))
        if val(ref,'customMarkFollows') is not None and enabled(ref,'customMarkFollows'):
            rp=ref.getparent();mark=ref.getnext()
            if mark is None:
                nr=rp.getnext()
                mark=nr.find('w:t',NS) if nr is not None and nr.tag==Q('r') else None
            if mark is None or mark.tag!=Q('t') or not re.fullmatch(r'(?:[0-9]+|[①-⑳]|[*†‡])',mark.text or ''):
                raise NumberingError('自定义脚注标记与正文混排，需先确认精确字符范围。')
            symbol=mark.text;np=note.find('w:p',NS)
            if np is None:raise NumberingError('自定义脚注缺少正文。')
            first=next(((n,a,b,v) for n,a,b,v in visible_nodes(np)),None)
            if first and first[3].startswith(symbol):
                if first[3]!=symbol and not first[3][len(symbol):].startswith((' ','\t')):
                    raise NumberingError('自定义脚注标记与脚注正文不能安全分开。')
                replace_span(np,0,len(symbol))
            elif not np.findall('.//w:footnoteRef',NS):
                raise NumberingError('自定义脚注正文标记与引用不匹配。')
            mark.text='';ref.attrib.pop(Q('customMarkFollows'),None)
            changes.append({'kind':'footnote','note_id':ident,'action':'自定义脚注标记改为自动序号'})
            pkg.touch('word/document.xml');pkg.touch('word/footnotes.xml')
        if not note.findall('.//w:footnoteRef',NS):
            p=note.find('w:p',NS)
            if p is None:raise NumberingError('脚注正文为空或结构不受支持。')
            m=NOTE_PREFIX.match(text(p))
            first_run=next((r for r in p.findall('w:r',NS) if text(r)),None)
            if m and first_run is not None and (
                val(first_run.find('w:rPr/w:vertAlign',NS))=='superscript' or
                re.search(r'footnote\s*reference|脚注引用',val(first_run.find('w:rPr/w:rStyle',NS),default=''),re.I)
            ):
                replace_span(p,0,m.end())
            index=1 if p.find('w:pPr',NS) is not None else 0;p.insert(index,footnote_ref_run(ident,True))
            changes.append({'kind':'footnote','note_id':ident,'action':'恢复脚注正文自动编号标记'});pkg.touch('word/footnotes.xml')
    # Reverse offsets allow more than one note in the same body paragraph.
    next_id=max((int(i) for i in ids if i is not None),default=0)+1
    assigned=[]
    for rp,start,end,np,m in sorted(plans,key=lambda x:(pkg.index[x[0]],x[1])):
        assigned.append((rp,start,end,np,m,str(next_id)));next_id+=1
    for rp,start,end,np,m,ident in reversed(assigned):
        original=text(np);moved=copy.deepcopy(np)
        replace_span(moved,0,m.end());move_note_relationships(pkg,moved)
        index=1 if moved.find('w:pPr',NS) is not None else 0
        moved.insert(index,footnote_ref_run(ident,True));moved.insert(index+1,run(' '))
        note=item('footnote',id=ident);note.append(moved);notes.append(note)
        replace_span(rp,start,end,[footnote_ref_run(ident)])
        np.getparent().remove(np)
        changes.append({'kind':'footnote','reference_paragraph':pkg.index[rp],'note_paragraph':pkg.index[np],
                        'before':original,'note_id':ident,'action':'普通上标及说明文字改为 Word 原生脚注'})
    if assigned:pkg.touch('word/document.xml');pkg.touch('word/footnotes.xml')


def apply_plan(pkg,path):
    if path is None:return []
    def unique(pairs):
        result={}
        for k,v in pairs:
            if k in result:raise NumberingError('numbering-plan 字段重复：'+k)
            result[k]=v
        return result
    data=json.loads(Path(path).read_text(encoding='utf-8-sig'),object_pairs_hook=unique)
    if not isinstance(data,dict) or set(data)-{'source_sha256','headings','captions','footnotes'}:
        raise NumberingError('numbering-plan 存在未知字段。')
    digest='sha256:'+hashlib.sha256(pkg.source.read_bytes()).hexdigest()
    if data.get('source_sha256')!=digest:raise NumberingError('numbering-plan 不属于当前输入文件，请重新定位。')
    def paragraph(index,expected):
        if type(index) is not int or not 1<=index<=len(pkg.paragraphs):raise NumberingError('段落编号无效。')
        p=pkg.paragraphs[index-1]
        if text(p)!=expected:raise NumberingError('numbering-plan 文字锚点与当前文件不一致。')
        return p
    for kind in ('headings','captions','footnotes'):
        if not isinstance(data.get(kind,[]),list):raise NumberingError('numbering-plan.'+kind+' 必须为数组。')
    for row in data.get('headings',[]):
        if not isinstance(row,dict) or set(row)!={'paragraph','text','level'} or type(row['level']) is not int or not 1<=row['level']<=9:
            raise NumberingError('headings 项必须提供 paragraph/text/level（1—9）。')
        p=paragraph(row['paragraph'],row['text']);level=row['level']-1
        sid=next((val(s,'styleId') for s in pkg.style_map.values() if val(s.find('w:pPr/w:outlineLvl',NS))==str(level) and val(s,'type')=='paragraph'),None)
        if sid is None:raise NumberingError('当前文件缺少已确认的对应标题样式。')
        pp=ppr(p);node=pp.find('w:pStyle',NS)
        if node is None:node=item('pStyle');pp.insert(0,node)
        node.set(Q('val'),sid);pkg.touch('word/document.xml')
    for row in data.get('captions',[]):
        if not isinstance(row,dict) or set(row)!={'paragraph','text'}:raise NumberingError('captions 项只支持 paragraph/text。')
        p=paragraph(row['paragraph'],row['text'])
        sid=next((val(s,'styleId') for s in pkg.style_map.values() if re.search(r'caption|题注|图题|表题',val(s.find('w:name',NS),default=''),re.I) and val(s,'type')=='paragraph'),None)
        if sid is None:raise NumberingError('当前文件缺少已确认的题注样式。')
        pp=ppr(p);node=pp.find('w:pStyle',NS)
        if node is None:node=item('pStyle');pp.insert(0,node)
        node.set(Q('val'),sid);pkg.touch('word/document.xml')
    for row in data.get('footnotes',[]):
        required={'reference_paragraph','reference_text','start','end','note_paragraph','note_text'}
        if not isinstance(row,dict) or set(row)!=required:raise NumberingError('footnotes 映射字段不完整或有未知字段。')
        p=paragraph(row['reference_paragraph'],row['reference_text']);paragraph(row['note_paragraph'],row['note_text'])
        if type(row['start']) is not int or type(row['end']) is not int or not 0<=row['start']<row['end']<=len(text(p)):
            raise NumberingError('脚注引用文字范围无效。')
    return data.get('footnotes',[])


def audit(source:Path):
    result=audit_package(Package(source));result['input']=str(source);return result


def repair(source:Path,out:Path,plan:Path|None=None):
    source,out=Path(source),Path(out)
    if source.resolve()==out.resolve():raise NumberingError('输出文件不得覆盖输入文档。')
    pkg=Package(source);changes=[]
    # Original valid auto objects are intentionally left alone.
    before=audit_package(pkg)
    if before['status']=='passed' and plan is None:
        pkg.save(out);return {'status':'passed','input':str(source),'output':str(out),'change_count':0,'changes':[],'audit':before}
    explicit=apply_plan(pkg,plan)
    repair_headings(pkg,changes)
    repair_captions(pkg,changes)
    repair_footnotes(pkg,changes,explicit)
    # The paragraph index needs rebuilding after manual notes were moved.
    pkg.paragraphs=pkg.document.xpath('.//w:p[not(ancestor::w:del or ancestor::w:moveFrom)]',namespaces=NS)
    pkg.index={p:i for i,p in enumerate(pkg.paragraphs,1)}
    after=audit_package(pkg)
    if after['status']!='passed':
        return {'status':'blocked','input':str(source),'output':None,'change_count':0,'changes':[],
                'issues':after['issues'],'message':'仍有编号问题；本次修改未写出，需确认后重新修复。'}
    if pkg.changed:pkg.update_fields()
    pkg.save(out)
    from layout_invariant_guard import snapshot_docx, INVARIANTS
    before_guard=snapshot_docx(source);after_guard=snapshot_docx(out)
    frozen=INVARIANTS['automatic-numbering']
    damaged=[key for key in frozen if before_guard[key]!=after_guard[key]]
    if damaged:
        out.unlink(missing_ok=True)
        raise NumberingError('自动编号修复改变了非目标对象：'+', '.join(sorted(damaged)))
    final=audit(out)
    if final['status']!='passed':
        out.unlink(missing_ok=True);raise NumberingError('保存后自动编号复查失败，未交付输出文件。')
    return {'status':'passed','input':str(source),'output':str(out),'change_count':len(changes),'changes':changes,
            'scope':'automatic-numbering','protected_invariants':'passed','audit':final}


def preserve_heading_numbering(root,source_styles:bytes,numbering:bytes|None):
    """Materialize valid inherited heading numPr before template style replacement.

    Uses a lightweight Package view; it does not change style definitions, numbering
    IDs, list text, field codes or body-text rules.
    """
    proxy=object.__new__(Package);proxy.document=root;proxy.styles=xml(source_styles)
    proxy.numbering=xml(numbering) if numbering else E.Element(Q('numbering'),nsmap={'w':W})
    proxy.style_map={val(n,'styleId'):n for n in proxy.styles.findall('w:style',NS)}
    count=0
    for p in root.findall('.//w:p',NS):
        if proxy.level(p) is None:continue
        num=proxy.num(p)
        if num and num['valid'] and p.find('w:pPr/w:numPr',NS) is None:
            set_number(p,num['num_id'],num['level']);count+=1
    return count
