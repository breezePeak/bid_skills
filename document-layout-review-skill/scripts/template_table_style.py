#!/usr/bin/env python3
"""Apply/audit TABLE APPEARANCE against the active template, never the input.

Only appearance is changed here. Widths, merges, text, captions, diagrams and the
package theme are preserved. A width-repair guard protects the normalized result.
Template style inheritance/conditional formatting are copied into a private,
flattened table style; template theme colors are resolved locally, not globally.
"""
from __future__ import annotations

import argparse
import copy
import colorsys
import json
import os
import re
import tempfile
from pathlib import Path
from lxml import etree as E
from numbering_policy import Doc, PolicyError, Q, NS, A, node, value, visible, parse, digest, file_digest, ordered_properties

APPEARANCE = {
    'tblPr': ('tblStyle', 'tblLook', 'tblBorders', 'shd', 'tblStyleRowBandSize', 'tblStyleColBandSize'),
    'tblPrEx': ('tblBorders', 'shd'),
    'trPr': ('cnfStyle',),
    'tcPr': ('tcBorders', 'shd', 'cnfStyle'),
    'pPr': ('shd', 'pBdr', 'cnfStyle'),
    'rPr': ('color', 'shd', 'highlight'),
}
NO_SHADING = lambda: node('shd', val='nil')
NO_COLOR = lambda: node('color', val='auto')
NO_HIGHLIGHT = lambda: node('highlight', val='none')


def signature(el):
    """Expanded XML names; ignore namespace prefixes and insignificant whitespace."""
    if el is None: return None
    return [el.tag, sorted(el.attrib.items()), (el.text or '').strip(), [signature(c) for c in el]]


def merge_props(dst, src):
    if src is None: return
    for child in src:
        old = dst.find(child.tag)
        if old is None: dst.append(copy.deepcopy(child))
        elif len(child) and len(old): merge_props(old, child)
        else: dst.replace(old, copy.deepcopy(child))


def child_props(parent, tag):
    result = parent.find(Q(tag))
    if result is None:
        result = node(tag); parent.insert(0, result)
    return result


def replace_props(parent, tag, names, expected):
    current = child_props(parent, tag)
    for name in names:
        for old in list(current.findall(Q(name))): current.remove(old)
        new = expected.find(Q(name)) if expected is not None else None
        if new is not None: current.append(copy.deepcopy(new))
    ordered_properties(current)


def own_nodes(tbl, local):
    for n in tbl.iter(Q(local)):
        nearest = next((a for a in n.iterancestors() if a.tag == Q('tbl')), None)
        if nearest is tbl and not n.xpath('ancestor::w:drawing|ancestor::w:pict|ancestor::w:txbxContent', namespaces=NS):
            yield n


def header_key(tbl):
    row = tbl.find(Q('tr'))
    return tuple(re.sub(r'[\s\u200b\ufeff]+', '', visible(c)) for c in row.findall(Q('tc'))) if row is not None else ()


def table_sid(doc, tbl):
    sid = value(tbl.find('w:tblPr/w:tblStyle', NS))
    if sid: return sid
    return next((value(s, 'styleId') for s in doc.styles.findall(Q('style'))
                 if value(s, 'type') == 'table' and value(s, 'default', '0') not in {'0','false','off'}), None)


def theme_colors(doc):
    colors = {}
    # Locate the main-document theme by relationship, not by an arbitrary file.
    relname = 'word/_rels/document.xml.rels'
    target = None
    if relname in doc.files:
        import posixpath
        for rel in parse(doc.files[relname]):
            if rel.get('Type', '').endswith('/theme') and rel.get('TargetMode') != 'External':
                target = posixpath.normpath(posixpath.join('word', rel.get('Target', ''))).lstrip('/')
    if target is None and 'word/theme/theme1.xml' in doc.files: target = 'word/theme/theme1.xml'
    if target in doc.files:
        root = parse(doc.files[target])
        for scheme in root.iter('{' + A + '}clrScheme'):
            for entry in scheme:
                if len(entry):
                    c = entry[0]; rgb = c.get('lastClr') if E.QName(c).localname == 'sysClr' else c.get('val')
                    if rgb and re.fullmatch('[0-9A-Fa-f]{6}', rgb): colors[E.QName(entry).localname] = rgb.upper()
    mapping = {'background1':'lt1', 'text1':'dk1', 'background2':'lt2', 'text2':'dk2'}
    if 'word/settings.xml' in doc.files:
        cm = parse(doc.files['word/settings.xml']).find(Q('clrSchemeMapping'))
        if cm is not None:
            for k, v in cm.attrib.items(): mapping[E.QName(k).localname] = v
    aliases = {'light1':'lt1','dark1':'dk1','light2':'lt2','dark2':'dk2',
               'hyperlink':'hlink','followedHyperlink':'folHlink','bg1':'lt1','t1':'dk1','bg2':'lt2','t2':'dk2'}
    for name, dest in {**aliases, **mapping}.items():
        if dest in colors: colors[name] = colors[dest]
    for name, slot in {'background1':'bg1','text1':'t1','background2':'bg2','text2':'t2'}.items():
        if slot in colors: colors[name] = colors[slot]
    return colors


def resolved_colors(el, colors):
    result = copy.deepcopy(el)
    for n in result.iter():
        for theme_attr, direct_attr, tint_attr, shade_attr in (
            ('themeColor', 'val' if n.tag == Q('color') else 'color', 'themeTint', 'themeShade'),
            ('themeFill', 'fill', 'themeFillTint', 'themeFillShade')):
            theme = value(n, theme_attr)
            if not theme: continue
            rgb = colors.get(theme)
            if rgb is None:
                raise PolicyError('template-theme-color-missing', '模板主题颜色无法解析，不能用待修文档的主题色代替。', color=theme)
            channels = [int(rgb[i:i+2], 16) / 255 for i in (0, 2, 4)]
            hue, lightness, saturation = colorsys.rgb_to_hls(*channels)
            # WordprocessingML tint is the retained luminance fraction; 00 is
            # white and FF is the original. Tint takes precedence over shade.
            if value(n, tint_attr) is not None:
                factor = int(value(n, tint_attr), 16) / 255
                lightness = lightness * factor + (1 - factor)
            elif value(n, shade_attr) is not None:
                lightness *= int(value(n, shade_attr), 16) / 255
            channels = [round(v * 255) for v in colorsys.hls_to_rgb(hue, lightness, saturation)]
            n.set(Q(direct_attr), ''.join(f'{v:02X}' for v in channels))
            for key in (theme_attr, tint_attr, shade_attr): n.attrib.pop(Q(key), None)
    return result


def style_flat(doc, sid, colors):
    if sid is None:
        return node('style', type='table', styleId='none')
    if sid not in doc.style_map or value(doc.style_map[sid], 'type') != 'table':
        raise PolicyError('template-table-style-missing', '生效模板缺少指定表格样式。', style_id=sid)
    chain = doc.chain(sid)
    if any(value(s.find(Q('basedOn'))) not in (None, *doc.style_map) for s in chain):
        raise PolicyError('template-style-parent-missing', '模板表格样式的继承来源缺失。')
    result = node('style', type='table', styleId=sid)
    for s in chain:
        for child in s:
            if child.tag in {Q('tblPr'),Q('trPr'),Q('tcPr'),Q('pPr'),Q('rPr')}:
                dst = result.find(child.tag)
                if dst is None: result.append(copy.deepcopy(child))
                else: merge_props(dst, child)
            elif child.tag == Q('tblStylePr'):
                key = value(child, 'type')
                dst = next((x for x in result.findall(Q('tblStylePr')) if value(x,'type') == key), None)
                if dst is None: result.append(copy.deepcopy(child))
                else: merge_props(dst, child)
    result.insert(0, node('name', val=value(doc.style_map[sid].find(Q('name')), default=sid)))
    return resolved_colors(result, colors)


def look_flags(look):
    bits = {'firstRow':0x20,'lastRow':0x40,'firstColumn':0x80,'lastColumn':0x100,'noHBand':0x200,'noVBand':0x400}
    mask = int(value(look, 'val', '0000'), 16)
    return {key: value(look,key,'1' if mask & bit else '0') not in {'0','false','off'} for key,bit in bits.items()}


def active_conditions(flat, look, row, col, rows, cols):
    f = look_flags(look); result = ['wholeTable']
    def band(tag):
        n = flat.find('w:tblPr/w:' + tag, NS)
        return max(1,int(value(n, default='1')))
    if not f['noVBand']:
        offset = col - int(f['firstColumn'])
        if offset >= 0: result.append('band1Vert' if offset // band('tblStyleColBandSize') % 2 == 0 else 'band2Vert')
    if not f['noHBand']:
        offset = row - int(f['firstRow'])
        if offset >= 0: result.append('band1Horz' if offset // band('tblStyleRowBandSize') % 2 == 0 else 'band2Horz')
    if row == 0 and f['firstRow']: result.append('firstRow')
    if row == rows-1 and f['lastRow']: result.append('lastRow')
    if col == 0 and f['firstColumn']: result.append('firstCol')
    if col == cols-1 and f['lastColumn']: result.append('lastCol')
    for label, yes in [('nwCell', row==0 and col==0 and f['firstRow'] and f['firstColumn']),
                       ('neCell', row==0 and col==cols-1 and f['firstRow'] and f['lastColumn']),
                       ('swCell', row==rows-1 and col==0 and f['lastRow'] and f['firstColumn']),
                       ('seCell', row==rows-1 and col==cols-1 and f['lastRow'] and f['lastColumn'])]:
        if yes: result.append(label)
    return result


def conditional_props(flat, look, kind, row, col, rows, cols):
    result = node(kind); merge_props(result, flat.find(Q(kind)))
    for typ in active_conditions(flat, look, row, col, rows, cols):
        for cond in flat.findall(Q('tblStylePr')):
            if value(cond, 'type') == typ: merge_props(result, cond.find(Q(kind)))
    return result


class TemplateTables:
    def __init__(self, template, style_json=None):
        self.doc = Doc(Path(template)); self.colors = theme_colors(self.doc)
        self.samples = list(self.doc.document.iter(Q('tbl')))
        self.json = json.loads(Path(style_json).read_text(encoding='utf-8-sig')) if style_json else {}
        declared = self.json.get('template', {}).get('canonical_template_sha256')
        if declared is not None and declared != file_digest(template):
            raise PolicyError('template-profile-stale', '样式约定不属于当前模板，须重新解析；不能拿旧 JSON 覆盖新模板。')
        self.canonical = str(self.json.get('table', {}).get('style_id') or '') or None
        self.profile_hash = file_digest(style_json) if style_json else None

    def choose(self, tbl):
        eligible = [s for s in self.samples if table_sid(self.doc,s) == self.canonical] if self.canonical else self.samples
        matching = [s for s in eligible if header_key(s) == header_key(tbl) and header_key(tbl)]
        choices = matching or eligible
        if not choices:
            if self.canonical:
                sample = node('tbl'); pr = node('tblPr'); pr.append(node('tblStyle',val=self.canonical)); sample.append(pr)
                choices = [sample]
            else:
                raise PolicyError('template-table-example-missing', '模板中未找到可确定的表格标准；须从模板补充表格样式约定，不能保留输入格式当作通过。')
        contracts = [self.contract(s) for s in choices]
        # Different content/widths do not constitute a style conflict.
        if len({json.dumps(c['key'],ensure_ascii=False,sort_keys=True) for c in contracts}) != 1:
            raise PolicyError('template-table-style-ambiguous', '模板有多套不同表格外观；请在本次样式约定中确定 table.style_id 或对应表头样表，不能按出现次数猜。')
        return contracts[0]

    def contract(self, sample):
        sid = self.canonical or table_sid(self.doc, sample)
        flat = style_flat(self.doc,sid,self.colors)
        look = sample.find('w:tblPr/w:tblLook',NS)
        if look is None: look = node('tblLook',val='0000',firstRow='0',lastRow='0',firstColumn='0',lastColumn='0',noHBand='1',noVBand='1')
        tablepr = node('tblPr')
        for source in (flat.find(Q('tblPr')),sample.find(Q('tblPr'))):
            if source is not None:
                for name in ('tblBorders','shd','tblStyleRowBandSize','tblStyleColBandSize'):
                    child = source.find(Q(name))
                    if child is not None:
                        old = tablepr.find(Q(name))
                        if old is not None: tablepr.remove(old)
                        tablepr.append(resolved_colors(child,self.colors))
        # Missing shading is a real template state, not permission to keep old fill.
        if tablepr.find(Q('shd')) is None: tablepr.append(NO_SHADING())
        # An explicit approved JSON border contract is used when Word stores the
        # template's borders in its sample, or when there is no physical sample.
        b = self.json.get('table',{}).get('borders',{})
        if b.get('all') is not None:
            old = tablepr.find(Q('tblBorders'))
            if old is not None: tablepr.remove(old)
            borders = node('tblBorders')
            for edge in ('top','left','bottom','right','insideH','insideV'):
                borders.append(node(edge,val=b['all'],sz=b.get('size_eighth_points',4),color=b.get('color','auto')))
            tablepr.append(borders)
        if tablepr.find(Q('tblBorders')) is None:
            borders = node('tblBorders')
            for edge in ('top','left','bottom','right','insideH','insideV'): borders.append(node(edge,val='nil'))
            tablepr.append(borders)
        tablepr.append(copy.deepcopy(look))
        rows = sample.findall(Q('tr'))
        patterns = [[self.cell_pattern(tc) for tc in row.findall(Q('tc'))] for row in rows]
        # A uniform header/body appearance can apply to any number of columns.
        compact = []
        for group in (patterns[:1], patterns[1:]):
            sigs = {json.dumps(p,sort_keys=True) for row in group for p in row}
            compact.append(sorted(sigs))
        key = [signature(flat), signature(tablepr), compact]
        token = digest(json.dumps(key,ensure_ascii=False,sort_keys=True).encode()).split(':')[1][:20]
        flat.set(Q('styleId'),'DLRT_'+token);flat.find(Q('name')).set(Q('val'),'模板表格 '+token)
        return {'style':flat,'props':tablepr,'sample':sample,'patterns':patterns,'key':key,
                'style_id':value(flat,'styleId'),'source_style_id':sid,
                'sample_index':self.samples.index(sample)+1 if sample in self.samples else None}

    def cell_pattern(self, tc):
        result = {}
        tcpr = tc.find(Q('tcPr'))
        for kind,names,parent in [('tcPr',('shd','tcBorders'),tcpr),('tblPrEx',('shd','tblBorders'),tc.getparent().find(Q('tblPrEx')))]:
            result[kind] = {}
            for name in names:
                el = parent.find(Q(name)) if parent is not None else None
                if el is not None: result[kind][name] = E.tostring(resolved_colors(el,self.colors)).decode()
        paras = [p for p in tc.findall(Q('p')) if visible(p).strip()] or tc.findall(Q('p'))
        per_p=[];per_r=[]
        for p in paras:
            pp = node('pPr')
            for st in self.doc.chain(self.doc.style_id(p)): merge_props(pp,st.find(Q('pPr')))
            merge_props(pp,p.find(Q('pPr')))
            per_p.append({k:E.tostring(resolved_colors(pp.find(Q(k)),self.colors)).decode() for k in ('shd','pBdr') if pp.find(Q(k)) is not None})
            runs = list(p.findall(Q('r')))
            if not runs: runs = [node('r')]
            for r in runs:
                rp = node('rPr')
                for st in self.doc.chain(self.doc.style_id(p)): merge_props(rp,st.find(Q('rPr')))
                cs = value(r.find('w:rPr/w:rStyle',NS))
                for st in self.doc.chain(cs): merge_props(rp,st.find(Q('rPr')))
                merge_props(rp,r.find(Q('rPr')))
                per_r.append({k:E.tostring(resolved_colors(rp.find(Q(k)),self.colors)).decode() for k in ('color','shd','highlight') if rp.find(Q(k)) is not None})
        for kind,items in [('pPr',per_p),('rPr',per_r)]:
            unique = {json.dumps(i,sort_keys=True) for i in items}
            if len(unique)>1: raise PolicyError('template-cell-appearance-mixed', '模板单元格内有多种文字颜色/底纹，需先确定通用表格文字约定，不能从旧文档猜。')
            result[kind] = items[0] if items else {}
        return result

    def pattern(self, contract, target, row, col, row_count, col_count):
        patterns=contract['patterns']
        if not patterns: return {k:{} for k in ('tcPr','tblPrEx','pPr','rPr')}
        exact=header_key(target)==header_key(contract['sample'])
        candidates=patterns[:1] if row==0 else patterns[1:]
        if not candidates: candidates=patterns[:1]
        if exact and all(len(r)==col_count for r in candidates): values=[r[col] for r in candidates]
        else: values=[p for r in candidates for p in r]
        signatures={json.dumps(p,sort_keys=True) for p in values}
        if len(signatures)!=1:
            raise PolicyError('template-region-ambiguous', '模板不同列或数据行的直接外观不同，无法自动映射到当前表格；需确认对应样表，不按旧外观放行。')
        return values[0]


def project(doc):
    result=[]
    for ti,tbl in enumerate(doc.document.iter(Q('tbl')),1):
        rows=[]
        for kind,names in APPEARANCE.items():
            if kind=='tblPr': elements=[tbl.find(Q(kind))]
            else: elements=list(own_nodes(tbl,kind))
            for idx,prop in enumerate(elements):
                if prop is not None:
                    picked=[signature(prop.find(Q(k))) for k in names]
                    if any(x is not None for x in picked):rows.append((kind,idx,picked))
        sid=table_sid(doc,tbl)
        chain=[signature(s) for s in doc.chain(sid)]
        result.append({'table':ti,'props':rows,'style':chain})
    return result


def apply(doc, standard):
    records=[]
    for ti,tbl in enumerate(doc.document.iter(Q('tbl')),1):
        contract=standard.choose(tbl)
        sid=contract['style_id'];new=contract['style']
        existing=doc.style_map.get(sid)
        if existing is None or signature(existing)!=signature(new):
            if existing is not None: doc.styles.remove(existing)
            doc.styles.append(copy.deepcopy(new));doc.changed.add('word/styles.xml');doc.refresh()
        expected=copy.deepcopy(contract['props']);expected.insert(0,node('tblStyle',val=sid))
        replace_props(tbl,'tblPr',APPEARANCE['tblPr'],expected)
        rows=tbl.findall(Q('tr'))
        for ri,row in enumerate(rows):
            # Original row exceptions and cnfStyle must not reactivate old fills.
            for prtag in ('tblPrEx','trPr'):
                pr=row.find(Q(prtag))
                if pr is not None:
                    for name in APPEARANCE[prtag]:
                        for old in list(pr.findall(Q(name))):pr.remove(old)
            cells=row.findall(Q('tc'))
            for ci,cell in enumerate(cells):
                pat=standard.pattern(contract,tbl,ri,ci,len(rows),len(cells))
                cp=node('tcPr')
                for payload in pat['tcPr'].values():cp.append(parse(payload.encode()))
                # Use cloned style for conditional cell fill, plus exact source
                # direct formatting; never leave original tcPr shading in place.
                replace_props(cell,'tcPr',APPEARANCE['tcPr'],cp)
                for key,payload in pat['tblPrEx'].items():
                    ex=child_props(row,'tblPrEx');old=ex.find(Q(key))
                    if old is not None:ex.remove(old)
                    ex.append(parse(payload.encode()))
                for p in cell.findall(Q('p')):
                    pp=node('pPr')
                    merge_props(pp,standard.doc.styles.find('w:docDefaults/w:pPrDefault/w:pPr',NS))
                    merge_props(pp,conditional_props(new,expected.find(Q('tblLook')),'pPr',ri,ci,len(rows),len(cells)))
                    for key,payload in pat['pPr'].items():
                        old=pp.find(Q(key))
                        if old is not None:pp.remove(old)
                        pp.append(parse(payload.encode()))
                    pp=resolved_colors(pp,standard.colors)
                    if pp.find(Q('shd')) is None:pp.append(NO_SHADING())
                    if pp.find(Q('pBdr')) is None:
                        pb=node('pBdr')
                        for edge in ('top','left','bottom','right','between','bar'):pb.append(node(edge,val='nil'))
                        pp.append(pb)
                    replace_props(p,'pPr',APPEARANCE['pPr'],pp)
                    base=node('rPr')
                    merge_props(base,standard.doc.styles.find('w:docDefaults/w:rPrDefault/w:rPr',NS))
                    merge_props(base,conditional_props(new,expected.find(Q('tblLook')),'rPr',ri,ci,len(rows),len(cells)))
                    for key,payload in pat['rPr'].items():
                        old=base.find(Q(key))
                        if old is not None:base.remove(old)
                        base.append(parse(payload.encode()))
                    base=resolved_colors(base,standard.colors)
                    for name,factory in [('color',NO_COLOR),('shd',NO_SHADING),('highlight',NO_HIGHLIGHT)]:
                        if base.find(Q(name)) is None:base.append(factory())
                    # Field-result runs are included; drawing/textbox runs are not.
                    for r in p.iter(Q('r')):
                        if r.xpath('ancestor::w:drawing|ancestor::w:pict|ancestor::w:txbxContent',namespaces=NS):continue
                        replace_props(r,'rPr',APPEARANCE['rPr'],base)
        doc.changed.add('word/document.xml')
        records.append({'table':ti,'template_style_id':contract['source_style_id'],
                        'template_sample':contract['sample_index'],'applied_style_id':sid})
    return records


def audit(source,template,style_json=None):
    doc=Doc(Path(source));before=project(doc)
    standard=TemplateTables(template,style_json)
    records=apply(doc,standard);expected=project(doc)
    issues=[{'severity':'error','code':'table-template-appearance-mismatch','table':a['table'],
             'message':'表格外观不符合生效模板；原有底色、边框或继承颜色不能作为通过依据。'} for a,b in zip(before,expected) if a!=b]
    return {'status':'failed' if issues else 'passed','input':str(source),'input_sha256':file_digest(source),
            'template':str(template),'template_sha256':file_digest(template),'template_style_sha256':standard.profile_hash,
            'tables':records,'issues':issues,'visual_review_required':True}


def repair(source,out,template,style_json=None):
    source,out,template=map(Path,(source,out,template))
    if out.resolve() in {source.resolve(),template.resolve()}:raise PolicyError('in-place-write','不得覆盖原文档或模板。')
    doc=Doc(source);before=project(doc);standard=TemplateTables(template,style_json)
    records=apply(doc,standard);after=project(doc)
    changes=[{**record,'action':'按模板统一表格底色、边框及文字颜色（含无底色状态）'} for record,a,b in zip(records,before,after) if a!=b]
    out.parent.mkdir(parents=True,exist_ok=True)
    with tempfile.TemporaryDirectory(dir=out.parent) as work:
        candidate=Path(work)/'candidate.docx';doc.save(candidate)
        from layout_invariant_guard import snapshot_docx,changed_invariants
        damage=changed_invariants(snapshot_docx(source),snapshot_docx(candidate),'template-table-style')
        if damage:raise PolicyError('template-style-collateral-edit','模板表格格式修复改变了内容、尺寸或非目标对象。',invariants=damage)
        checked=audit(candidate,template,style_json)
        if checked['status']!='passed':raise PolicyError('template-style-recheck-failed','保存后模板一致性检查失败。',issues=checked['issues'])
        os.replace(candidate,out)
    return {'status':'passed','input':str(source),'output':str(out),'output_sha256':file_digest(out),
            'template':str(template),'template_sha256':file_digest(template),'template_style_sha256':standard.profile_hash,
            'scope':'template-table-style','change_count':len(changes),'changes':changes,'audit':checked}


def main():
    ap=argparse.ArgumentParser(description='以模板统一/检查表格外观；不保留与模板冲突的旧底色')
    ap.add_argument('action',choices=('repair','audit'));ap.add_argument('input',type=Path)
    ap.add_argument('--template',type=Path,required=True);ap.add_argument('--template-style-json',type=Path)
    ap.add_argument('--out',type=Path);ap.add_argument('--json-out',type=Path)
    a=ap.parse_args()
    try:
        if a.action=='repair' and a.out is None:raise PolicyError('output-required','修复必须指定 --out。')
        result=repair(a.input,a.out,a.template,a.template_style_json) if a.action=='repair' else audit(a.input,a.template,a.template_style_json)
    except Exception as exc:
        result={'status':'blocked','output':None,'issues':[exc.as_issue() if isinstance(exc,PolicyError) else {'severity':'error','code':'template-table-style-error','message':str(exc)}]}
    text=json.dumps(result,ensure_ascii=False,indent=2)
    if a.json_out and a.json_out.resolve() in {p.resolve() for p in (a.input,a.template,a.out,a.template_style_json) if p is not None}:
        print(json.dumps({'status':'blocked','issues':[{'code':'report-output-collision','message':'报告路径不能覆盖文档、模板或规则。'}]},ensure_ascii=False));return 2
    if a.json_out:a.json_out.parent.mkdir(parents=True,exist_ok=True);a.json_out.write_text(text+'\n',encoding='utf-8')
    print(text);return 0 if result['status']=='passed' else 2


if __name__=='__main__':raise SystemExit(main())
