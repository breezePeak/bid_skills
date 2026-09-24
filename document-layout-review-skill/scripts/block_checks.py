"""Read-only checks of the selected block; no whole-document repair or Office call."""
from __future__ import annotations

import copy
import io
import json
import re
import zipfile
from collections import Counter
from pathlib import Path
from lxml import etree as E

from block_progress import BlockError, DOC, NS, Q, package, read_json, sha
from template_format_contract import (StyleResolver, paragraph_style, own_runs,
    expected_style_id, template_role_ids, audit_paragraph, value_key, mismatches)
import content_integrity as content


def fragment(parts, part, start, end):
    """Feed existing content rules a tiny, in-memory package, not the whole book."""
    root = content.parse(parts[part])
    parent = root.find(Q('body')) if part == DOC else root
    children = list(parent)
    chosen = children[start:end]
    for child in children:
        parent.remove(child)
    parent.extend(chosen)
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_STORED) as archive:
        archive.writestr('word/styles.xml', parts['word/styles.xml'])
        archive.writestr(part, E.tostring(root))
        # Include only notes referenced in the selected block.
        for kind in ('footnote', 'endnote'):
            name = f'word/{kind}s.xml'
            if name == part or name not in parts:
                continue
            wanted = {n.get(Q('id')) for n in root.iter(Q(kind+'Reference'))}
            if wanted:
                notes = content.parse(parts[name])
                for child in list(notes):
                    if child.get(Q('id')) not in wanted:
                        notes.remove(child)
                archive.writestr(name, E.tostring(notes))
    out.seek(0)
    return out


def field_structure(nodes):
    stack = []
    for root in nodes:
        for node in root.iter():
            if node.tag == Q('fldChar'):
                kind = node.get(Q('fldCharType'))
                if kind == 'begin':
                    stack.append(False)
                elif kind == 'separate':
                    if not stack or stack[-1]:
                        raise BlockError('当前块域分隔符没有唯一的开始边界。')
                    stack[-1] = True
                elif kind == 'end':
                    if not stack:
                        raise BlockError('当前块域结束符没有对应开始边界。')
                    stack.pop()
                if node.get(Q('fldLock')) in {'true', '1', 'on'}:
                    raise BlockError('不能锁定域来掩盖编号问题。')
            elif node.tag == Q('fldSimple'):
                if not (node.get(Q('instr')) or '').strip():
                    raise BlockError('当前块存在空域指令。')
                if node.get(Q('fldLock')) in {'true', '1', 'on'}:
                    raise BlockError('不能锁定域来掩盖编号问题。')
    if stack:
        raise BlockError('当前块包含未闭合的域；先修复边界，不得标记完成。')


def _content_issues(state, block, proposal, delta):
    before, _, _ = package(state['current'])
    after, _, _ = package(proposal)
    part = block.get('part', DOC)
    a, b = block['start'], block['end']
    old = content.project(fragment(before, part, a, b))
    new = content.project(fragment(after, part, a, b+delta))
    plan = read_json(state['content_plan']) if state.get('content_plan') else {}
    if plan and plan.get('source_sha256') != state['source_sha256']:
        raise BlockError('内容保护计划不属于原始文档。')
    issues = []
    for key in ('body', 'protected_literals', 'tracked_deletions'):
        issues.extend(content.compare_sequence(old[key], new[key], scope=block['id']+'/'+key))
    for name in sorted(set(old['other_stories']) | set(new['other_stories'])):
        issues.extend(content.compare_sequence(old['other_stories'].get(name, []),
                      new['other_stories'].get(name, []), scope=block['id']+'/'+name))
    for kind in ('figure', 'table'):
        approved = Counter(content.canonical_text(r['title']) for r in plan.get('allowed_caption_insertions', [])
                           if r.get('id') in block.get('objects', []) and r.get('kind') == kind)
        i, kept = 0, []
        for title in new['captions'][kind]:
            if i < len(old['captions'][kind]) and old['captions'][kind][i] == title:
                kept.append(title); i += 1
            elif approved[title] > 0:
                approved[title] -= 1
            else:
                kept.append(title)
        issues.extend(content.compare_sequence(old['captions'][kind], kept, scope=block['id']+'/'+kind))
    # Native note conversion is rare. Use the original-bound validator for that
    # structural operation only; never suppress a local business-text difference.
    if issues and plan.get('allowed_footnote_moves'):
        result = content.audit(state['source'], proposal, plan)
        if result['status'] == 'passed':
            return []
    return issues


def _font_key(value):
    value = value.strip().lstrip('@').casefold()
    return {'黑体':'simhei', '宋体':'simsun'}.get(value, value)


def resolved_fonts(props, parts):
    fonts = props.find(Q('rFonts'))
    if fonts is None:
        return {}
    out = {E.QName(k).localname:v for k,v in fonts.attrib.items()}
    theme = content.parse(parts['word/theme/theme1.xml']) if 'word/theme/theme1.xml' in parts else None
    settings = content.parse(parts['word/settings.xml']) if 'word/settings.xml' in parts else None
    lang = settings.find(Q('themeFontLang')) if settings is not None else None
    east = lang.get(Q('eastAsia')) if lang is not None else None
    script = {'zh-CN':'Hans','zh-SG':'Hans','zh-TW':'Hant','zh-HK':'Hant','ja-JP':'Jpan','ko-KR':'Hang'}.get(east)
    A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    result = {}
    for slot in ('ascii','hAnsi','eastAsia','cs'):
        themed = out.get(slot+'Theme') or (out.get('cstheme') if slot=='cs' else None)
        face = out.get(slot)
        if themed:
            branch = theme.find('.//'+A+('majorFont' if themed.startswith('major') else 'minorFont')) if theme is not None else None
            node = branch.find(A+({'eastAsia':'ea','cs':'cs'}.get(slot,'latin'))) if branch is not None else None
            face = node.get('typeface') if node is not None else None
            if not face and branch is not None and slot=='eastAsia' and script:
                node = branch.find(A+f"font[@script='{script}']")
                face = node.get('typeface') if node is not None else None
            if not face:
                # An unresolved identical theme can be compared, not guessed.
                face = 'theme:'+themed
        if face:
            result[slot] = _font_key(face)
    return result


def _styles(state, nodes, parts, part):
    settings = state.get('settings', {})
    path = settings.get('template')
    if not path:
        return [], 0  # cursor-only API, never evidence of template conformance
    template_parts, _, _ = package(path)
    expected = StyleResolver(template_parts['word/styles.xml'])
    current = StyleResolver(parts['word/styles.xml'])
    profile = settings.get('template_style_json')
    if profile and not Path(profile).is_file():
        raise BlockError('模板约定文件缺失，不能通过块内检查。')
    roles = template_role_ids(profile)
    rules_path = settings.get('text_rules')
    rules = read_json(rules_path).get('body', {}) if rules_path else {}
    issues, count = [], 0
    for root in nodes:
        for p in root.iter(Q('p')):
            if p.xpath('ancestor::w:drawing|ancestor::w:pict|ancestor::w:txbxContent',namespaces=NS):
                continue
            if not list(own_runs(p)):
                continue
            sid = paragraph_style(p,current)
            wanted = expected_style_id(sid,expected,roles)
            if wanted is None:
                issues.append({'code':'block-unmapped-style','style_id':sid}); continue
            count += 1
            in_cell = bool(p.xpath('ancestor::w:tc',namespaces=NS))
            # Conditional table appearance is audited by the table checker. Here
            # enforce its explicit paragraph text role, not the body-font exception.
            body = not in_cell and part==DOC and (wanted==roles.get('body') or (not roles.get('body') and wanted==expected.default_paragraph))
            app=current.properties('pPr',sid,p.find(Q('pPr')))
            epp=expected.properties('pPr',wanted)
            basic=mismatches(app,epp,{'spacing','ind','jc'})
            if basic:
                issues.append({'code':'block-paragraph-format','paragraph':count,'conflicts':basic})
            if not in_cell:
                issues.extend(dict(code='block-extra-format',paragraph=count,**r) for r in audit_paragraph(p,wanted,expected,current))
            for ri,r in enumerate(own_runs(p),1):
                rp=r.find(Q('rPr')); csn=rp.find(Q('rStyle')) if rp is not None else None
                cs=csn.get(Q('val')) if csn is not None else None
                if cs and cs not in expected.styles:
                    issues.append({'code':'block-character-style','paragraph':count,'run':ri,'style_id':cs});continue
                arp=current.properties('rPr',sid,rp,character=cs)
                erp=expected.properties('rPr',wanted,character=cs)
                af,ef=resolved_fonts(arp,parts),resolved_fonts(erp,template_parts)
                if body:
                    ef.update({k:_font_key(v) for k,v in rules.get('fonts',{}).items()})
                banned={_font_key(v) for v in rules.get('forbidden_fonts',[])} if body else set()
                preserved={_font_key(v) for v in rules.get('preserve_fonts',[])}-banned if body else set()
                conflict={k:{'actual':af.get(k),'expected':v} for k,v in ef.items()
                          if af.get(k)!=v and af.get(k) not in preserved}
                if any(v in banned for v in af.values()):
                    conflict['forbidden_fonts']=sorted(banned & set(af.values()))
                if conflict:
                    issues.append({'code':'block-font','paragraph':count,'run':ri,'conflicts':conflict})
                for tag in ('sz','szCs','b','bCs','i','iCs'):
                    exp=value_key(erp,tag); actual=value_key(arp,tag)
                    if body and tag in {'sz','szCs'} and rules.get('size_pt') is not None:
                        exp=('explicit-size',str(int(rules['size_pt']*2)))
                        n=arp.find(Q(tag));actual=('explicit-size',n.get(Q('val')) if n is not None else None)
                    mode=rules.get('bold' if tag in {'b','bCs'} else 'italic','template') if body else 'template'
                    if tag in {'b','bCs','i','iCs'}:
                        if mode=='preserve':continue
                        if mode in {'require','forbid'}:exp=mode=='require'
                    if actual!=exp:
                        issues.append({'code':'block-run-format','paragraph':count,'run':ri,'property':tag,'actual':actual,'expected':exp})
    return issues, count


def check_local(state, block, proposal, delta):
    """Raise before mutation of state. A check receipt is not final page acceptance."""
    parts, _, body = package(proposal)
    part = block.get('part',DOC)
    parent = body if part==DOC else content.parse(parts[part])
    nodes = list(parent)[block['start']:block['end']+delta]
    field_structure(nodes)
    issues = _content_issues(state,block,proposal,delta)
    style_issues,count = _styles(state,nodes,parts,part)
    issues.extend(style_issues)
    if issues:
        error=BlockError('当前块检查失败，未采纳拟稿、未推进：'+json.dumps(issues[:12],ensure_ascii=False))
        error.issues=issues
        raise error
    return {'version':1,'status':'passed','block':block['id'],'proposal_sha256':sha(proposal),
            'checked_paragraphs':count,'content_checked':True,'field_structure_checked':True,
            'template_checked':bool(state.get('settings',{}).get('template')),
            'page_review':'pending-final-layout'}
