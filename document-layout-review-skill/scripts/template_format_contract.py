#!/usr/bin/env python3
"""Shared template property resolution for repair and audit.

This module does not classify business content, rewrite text, or repair diagrams.
Font/size and explicit body rules remain owned by text_rules.py. It closes the
previously unexamined direct/character-style formatting properties and copies
style dependencies instead of trusting the damaged document's definitions.
"""
from __future__ import annotations

import copy
import json
import zipfile
from pathlib import Path
from typing import Iterable
from lxml import etree as E

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}
Q = lambda tag: '{' + W + '}' + tag
FALSE = {'0', 'false', 'off'}
# These belong to text_rules/numbering/footnote handling, not the extra-format pass.
BASE_RUN = {'rFonts', 'sz', 'szCs', 'b', 'bCs', 'i', 'iCs', 'rStyle',
            'lang', 'noProof', 'vertAlign', 'vanish', 'specVanish', 'webHidden'}
EXTRA_RUN = {'color', 'u', 'highlight', 'shd', 'bdr', 'strike', 'dstrike',
             'caps', 'smallCaps', 'outline', 'shadow', 'emboss', 'imprint',
             'spacing', 'w', 'kern', 'position', 'effect', 'snapToGrid'}
EXTRA_PARAGRAPH = {'shd', 'pBdr', 'contextualSpacing', 'mirrorIndents',
                   'suppressAutoHyphens', 'suppressLineNumbers', 'autoSpaceDE',
                   'autoSpaceDN', 'textAlignment', 'textDirection', 'bidi',
                   'snapToGrid', 'adjustRightInd', 'wordWrap', 'widowControl'}
TOGGLES = {'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'outline',
           'shadow', 'emboss', 'imprint', 'vanish'}
BOOLS = TOGGLES | {'dstrike', 'contextualSpacing', 'mirrorIndents',
                  'suppressAutoHyphens', 'suppressLineNumbers', 'autoSpaceDE',
                  'autoSpaceDN', 'bidi', 'snapToGrid', 'adjustRightInd', 'wordWrap',
                  'widowControl'}
# Word defaults for properties that are on even in the absence of an element.
DEFAULT_ON = {'snapToGrid', 'autoSpaceDE', 'autoSpaceDN', 'wordWrap', 'widowControl'}
STORY = ('document', 'header', 'footer', 'footnotes', 'endnotes')


class FormatContractError(ValueError):
    """An unresolved baseline is a blocking issue, never a guessed format."""
    def __init__(self, code: str, message: str, **details):
        super().__init__(message)
        self.code, self.details = code, details

    def as_issue(self):
        return {'severity': 'error', 'code': self.code, 'message': str(self), **self.details}


def parse(data: bytes):
    return E.fromstring(data, E.XMLParser(resolve_entities=False, no_network=True))


def serialize(root):
    return E.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)


def signature(node):
    if node is None:
        return None
    return (node.tag, tuple(sorted(node.attrib.items())), (node.text or '').strip(),
            tuple(signature(child) for child in node))


def merge(base, extra, *, style=False):
    """Merge partial property attributes. Toggle semantics only apply to styles."""
    if extra is None:
        return
    for child in extra:
        if not isinstance(child.tag, str):
            continue
        name = E.QName(child).localname
        previous = base.find(child.tag)
        if style and name in TOGGLES:
            enabled = child.get(Q('val'), '1').lower() not in FALSE
            if not enabled:  # false in a style leaves the inherited value unchanged
                continue
            was = previous is not None and previous.get(Q('val'), '1').lower() not in FALSE
            replacement = E.Element(child.tag, {Q('val'): '0' if was else '1'})
        else:
            replacement = copy.deepcopy(child)
            if previous is not None:
                for attr, value in previous.attrib.items():
                    if attr not in replacement.attrib:
                        replacement.set(attr, value)
                if len(previous) and len(child):
                    replacement = copy.deepcopy(previous)
                    merge(replacement, child, style=False)
        if previous is None:
            base.append(replacement)
        else:
            base.replace(previous, replacement)


class StyleResolver:
    def __init__(self, styles: bytes):
        self.root = parse(styles)
        self.styles = {s.get(Q('styleId')): s for s in self.root.findall(Q('style'))}
        self.defaults = self.root.find(Q('docDefaults'))
        self.default_paragraph = next((sid for sid, s in self.styles.items()
                                       if s.get(Q('type')) == 'paragraph'
                                       and s.get(Q('default'), '0') not in FALSE), None)
        self._cache = {}

    def chain(self, sid: str | None):
        if sid is None:
            return []
        if sid in self._cache:
            return self._cache[sid]
        chain, seen, current = [], set(), sid
        while current:
            if current in seen:
                raise FormatContractError('template-style-cycle', '样式继承形成循环。', style_id=sid)
            if current not in self.styles:
                raise FormatContractError('template-style-dependency-missing', '样式或父样式缺失。', style_id=current)
            seen.add(current)
            item = self.styles[current]
            chain.append(item)
            parent = item.find(Q('basedOn'))
            current = parent.get(Q('val')) if parent is not None else None
        chain.reverse()
        self._cache[sid] = chain
        return chain

    def properties(self, kind: str, sid: str | None, direct=None,
                   character: str | None = None):
        result = E.Element(Q(kind))
        if self.defaults is not None:
            default = self.defaults.find(Q(kind + 'Default'))
            if default is not None:
                merge(result, default.find(Q(kind)))
        for style in self.chain(sid):
            merge(result, style.find(Q(kind)), style=True)
        if kind == 'rPr' and character:
            for style in self.chain(character):
                merge(result, style.find(Q(kind)), style=True)
        merge(result, direct)
        return result

    def dependency_ids(self, requested: Iterable[str]):
        result, pending = set(), list(requested)
        while pending:
            sid = pending.pop()
            if sid in result:
                continue
            # A link back or next=self is valid; a basedOn cycle is not.
            self.chain(sid)
            result.add(sid)
            for tag in ('basedOn', 'link', 'next'):
                ref = self.styles[sid].find(Q(tag))
                if ref is not None and ref.get(Q('val')):
                    pending.append(ref.get(Q('val')))
        return result


def copy_style_dependencies(target: bytes, template: bytes, requested: Iterable[str]):
    source = StyleResolver(template)
    target_root = parse(target)
    current = {s.get(Q('styleId')): s for s in target_root.findall(Q('style'))}
    needed = source.dependency_ids(requested)
    for sid in sorted(needed):
        replacement = copy.deepcopy(source.styles[sid])
        if sid in current:
            target_root.replace(current[sid], replacement)
        else:
            target_root.append(replacement)
    old = target_root.find(Q('docDefaults'))
    if old is not None:
        target_root.remove(old)
    if source.defaults is not None:
        target_root.insert(0, copy.deepcopy(source.defaults))
    return serialize(target_root)


def value_key(props, tag):
    node = props.find(Q(tag))
    if tag in BOOLS:
        return (tag in DEFAULT_ON) if node is None else node.get(Q('val'), '1').lower() not in FALSE
    if node is None:
        return {'color': 'auto', 'u': 'none', 'highlight': 'none', 'effect': 'none',
                'spacing': '0', 'position': '0', 'w': '100', 'kern': '0',
                'textAlignment': 'auto', 'textDirection': 'lrTb'}.get(tag)
    attrs = {E.QName(k).localname: v for k, v in node.attrib.items()}
    if tag == 'shd' and attrs.get('val', 'clear') == 'nil':
        return None
    if tag in {'u', 'highlight', 'effect'} and attrs.get('val') == 'none':
        return 'none'
    if tag == 'color' and attrs == {'val': 'auto'}:
        return 'auto'
    if tag in {'spacing', 'position', 'w', 'kern', 'textAlignment', 'textDirection'} and set(attrs) == {'val'}:
        return attrs['val']
    # Ignore a redundant auto underline color in the no-underline state.
    return signature(node)


def mismatches(actual, expected, names):
    return {name: {'actual': value_key(actual, name), 'expected': value_key(expected, name)}
            for name in sorted(names) if value_key(actual, name) != value_key(expected, name)}


def ensure(parent, tag):
    result = parent.find(Q(tag))
    if result is None:
        result = E.Element(Q(tag))
        parent.insert(0, result)
    return result


def replace_selected(parent, kind, names, expected):
    props = ensure(parent, kind)
    for name in names:
        for old in list(props.findall(Q(name))):
            props.remove(old)
        new = expected.find(Q(name))
        if new is not None:
            props.append(copy.deepcopy(new))
    # Keep schema order when a full bundle is available; no unrelated mutation.
    try:
        from numbering_policy import ordered_properties
    except ImportError:
        ordered_properties = None
    if ordered_properties:
        ordered_properties(props)


def paragraph_style(p, resolver):
    sid = p.find('w:pPr/w:pStyle', NS)
    return sid.get(Q('val')) if sid is not None else resolver.default_paragraph


def own_runs(p):
    for r in p.iter(Q('r')):
        if any(a.tag in {Q('drawing'), Q('pict'), Q('txbxContent')} for a in r.iterancestors()):
            continue
        owner = next((a for a in r.iterancestors() if a.tag == Q('p')), None)
        if owner is p and r.find(Q('t')) is not None:
            yield r


def repair_paragraph(p, sid: str, expected: StyleResolver, current: StyleResolver):
    """Set only extra properties, keeping body font rules and semantic markers."""
    changes = []
    actual_sid = paragraph_style(p, current)
    if sid not in expected.styles:
        raise FormatContractError('paragraph-template-style-missing', '段落尚未绑定到模板样式。', style_id=sid)
    epp = expected.properties('pPr', sid)
    pp = p.find(Q('pPr'))
    app = current.properties('pPr', actual_sid, pp)
    bad = mismatches(app, epp, EXTRA_PARAGRAPH)
    if bad:
        replace_selected(p, 'pPr', EXTRA_PARAGRAPH, epp)
        changes.append({'kind': 'paragraph', 'properties': sorted(bad)})
    # Paragraph-mark rPr is also a possible source of inherited-looking pollution.
    if pp is not None and pp.find(Q('rPr')) is not None:
        replace_selected(pp, 'rPr', EXTRA_RUN, expected.properties('rPr', sid))
    for index, r in enumerate(own_runs(p), 1):
        rp = r.find(Q('rPr'))
        cs_node = rp.find(Q('rStyle')) if rp is not None else None
        cs = cs_node.get(Q('val')) if cs_node is not None else None
        if cs and cs not in expected.styles:
            # Source-only character styles are never a template baseline.
            rp.remove(cs_node)
            changes.append({'kind': 'run', 'run': index, 'properties': ['non-template-character-style']})
            cs = None
        erp = expected.properties('rPr', sid, character=cs)
        arp = current.properties('rPr', actual_sid, rp, character=cs)
        bad = mismatches(arp, erp, EXTRA_RUN)
        if bad:
            replace_selected(r, 'rPr', EXTRA_RUN, erp)
            changes.append({'kind': 'run', 'run': index, 'properties': sorted(bad)})
    return changes


def audit_paragraph(p, sid: str, expected: StyleResolver, current: StyleResolver):
    issues = []
    actual_sid = paragraph_style(p, current)
    epp = expected.properties('pPr', sid)
    app = current.properties('pPr', actual_sid, p.find(Q('pPr')))
    bad = mismatches(app, epp, EXTRA_PARAGRAPH)
    if bad:
        issues.append({'kind': 'paragraph', 'conflicts': bad})
    pp=p.find(Q('pPr'))
    mark=pp.find(Q('rPr')) if pp is not None else None
    if mark is not None:
        erp=expected.properties('rPr',sid)
        arp=current.properties('rPr',actual_sid,mark)
        bad=mismatches(arp,erp,EXTRA_RUN)
        if bad:issues.append({'kind':'paragraph-mark','conflicts':bad})
    for index, r in enumerate(own_runs(p), 1):
        rp = r.find(Q('rPr'))
        cs_node = rp.find(Q('rStyle')) if rp is not None else None
        cs = cs_node.get(Q('val')) if cs_node is not None else None
        if cs and cs not in expected.styles:
            issues.append({'kind': 'run', 'run': index, 'conflicts': {'character_style': cs}})
            continue
        erp = expected.properties('rPr', sid, character=cs)
        arp = current.properties('rPr', actual_sid, rp, character=cs)
        bad = mismatches(arp, erp, EXTRA_RUN)
        if bad:
            issues.append({'kind': 'run', 'run': index, 'conflicts': bad})
    return issues



def template_role_ids(style_json=None):
    if style_json is None:return {}
    data=json.loads(Path(style_json).read_text(encoding='utf-8-sig'))
    roles={}
    for role,item in (data.get('semantic_styles') or data.get('semantic_roles') or {}).items():
        if not isinstance(item,dict):continue
        sid=item.get('source_style_id') or item.get('style_id')
        if sid is None and len(item.get('style_ids',[]))==1:sid=item['style_ids'][0]
        if sid is not None:roles[role]=str(sid)
    return roles


def expected_style_id(sid, expected, roles=None):
    """Map only the existing numbering core's private aliases, never source format."""
    if sid in expected.styles:return sid
    import re
    m=re.fullmatch(r'DLR(_|Next_)(heading[1-9]|caption)',sid or '')
    if m is None:return None
    role=m[2];source=(roles or {}).get(role)
    if source is None:
        choices=[]
        for ident,style in expected.styles.items():
            name=style.find(Q('name'));name=name.get(Q('val'),'') if name is not None else ''
            if role=='caption':matched=bool(re.search(r'caption|题注|图题|表题',name,re.I))
            else:
                level=expected.properties('pPr',ident).find(Q('outlineLvl'))
                matched=(level is not None and level.get(Q('val'))==str(int(role[-1])-1)) or bool(re.fullmatch(r'(?:heading|标题)\s*'+role[-1],name,re.I))
            if matched and style.get(Q('type'))=='paragraph':choices.append(ident)
        if len(choices)==1:source=choices[0]
    if source not in expected.styles:return None
    if m[1]=='Next_':
        nxt=expected.styles[source].find(Q('next'))
        source=nxt.get(Q('val')) if nxt is not None else None
    return source if source in expected.styles else None

def audit_extra_formats(template: Path, target: Path, style_json=None):
    """Every mapped story paragraph is checked; unknown styles are not a PASS."""
    issues, count = [], 0
    roles=template_role_ids(style_json)
    with zipfile.ZipFile(template) as tz, zipfile.ZipFile(target) as dz:
        expected = StyleResolver(tz.read('word/styles.xml'))
        current = StyleResolver(dz.read('word/styles.xml'))
        for part in dz.namelist():
            if not part.startswith('word/') or not part.endswith('.xml') or not Path(part).stem.startswith(STORY):
                continue
            root = parse(dz.read(part))
            for index, p in enumerate(root.iter(Q('p')), 1):
                if p.xpath('ancestor::w:drawing|ancestor::w:pict|ancestor::w:txbxContent', namespaces=NS):
                    continue  # separately required, hash-bound figure review
                if p.xpath('ancestor::w:tc', namespaces=NS):
                    continue  # conditional table formatting is checked by template_table_style
                if not list(own_runs(p)):
                    continue
                sid = expected_style_id(paragraph_style(p,current),expected,roles)
                if sid not in expected.styles:
                    issues.append({'severity': 'error', 'code': 'unmapped-paragraph-format',
                                   'part': part, 'paragraph': index, 'style_id': sid,
                                   'message': '尚未映射到模板样式，不能以未检查作为通过。'})
                    continue
                count += 1
                for issue in audit_paragraph(p, sid, expected, current):
                    issues.append({'severity': 'error', 'code': 'template-effective-format-mismatch',
                                   'part': part, 'paragraph': index, **issue})
    return {'status': 'failed' if issues else 'passed', 'checked_paragraphs': count,
            'issues': issues, 'delegated': ['body-font-size-emphasis', 'table-conditional-format', 'figure-visual-review']}


def main():
    import argparse
    ap=argparse.ArgumentParser(description='核验模板实际生效的补充段落/文字格式')
    ap.add_argument('template',type=Path);ap.add_argument('target',type=Path)
    ap.add_argument('--json-out',type=Path);ap.add_argument('--template-style-json',type=Path);a=ap.parse_args()
    try:result=audit_extra_formats(a.template,a.target,a.template_style_json)
    except Exception as exc:
        result={'status':'failed','issues':[exc.as_issue() if hasattr(exc,'as_issue') else
                {'severity':'error','code':'template-effective-format-error','message':str(exc)}]}
    payload=json.dumps(result,ensure_ascii=False,indent=2)
    if a.json_out:
        if a.json_out.resolve() in {a.template.resolve(),a.target.resolve()}:raise SystemExit('报告不能覆盖输入文件')
        a.json_out.parent.mkdir(parents=True,exist_ok=True);a.json_out.write_text(payload+'\n',encoding='utf-8')
    print(payload);return 0 if result['status']=='passed' else 2


if __name__=='__main__':raise SystemExit(main())
