"""Compare Office serialization by effective properties, never by raw style bytes.

Unknown properties remain significant. Kern and widowControl retain their real
values: this is not a blanket compatibility whitelist.
"""
from __future__ import annotations
import copy
from lxml import etree as E
from block_progress import package, Q, parse
from template_format_contract import StyleResolver, value_key, merge, BOOLS
from block_checks import resolved_fonts


def canonical(node):
    if node is None:return None
    local=E.QName(node).localname
    if local in {'pPr','rPr','tblPr','tcPr','trPr'}:
        keys={E.QName(c).localname for c in node}
        if local=='rPr':keys|={'kern','spacing','position','w','b','bCs','i','iCs'}
        if local=='pPr':keys|={'widowControl','jc','ind','spacing','keepNext','keepLines','pageBreakBefore'}
        result={}
        for key in sorted(keys):
            c=node.find(Q(key))
            if key in BOOLS or key in {'kern','spacing','position','w','sz','szCs','jc','ind','outlineLvl'}:
                result[key]=value_key(node,key)
            elif c is not None:result[key]=canonical(c)
        return result
    attrs=tuple(sorted(node.attrib.items()))
    if local in BOOLS:
        attrs=((Q('val'),'0' if node.get(Q('val'),'1') in {'0','false','off'} else '1'),)
    return (node.tag,attrs,(node.text or '').strip(),tuple(sorted((canonical(c) for c in node),key=repr)))


def style_signature(parts, sid):
    resolver=StyleResolver(parts['word/styles.xml'])
    if sid not in resolver.styles:return None
    style=resolver.styles[sid]
    result={'type':style.get(Q('type'))}
    for kind in ('pPr','rPr','tblPr','trPr','tcPr'):
        props=resolver.properties(kind,sid)
        # Theme-font identity is compared after resolution, not a fallback face.
        if kind=='rPr':
            result['fonts']=resolved_fonts(props,parts)
            font=props.find(Q('rFonts'))
            if font is not None:props.remove(font)
        result[kind]=canonical(props)
    conditions={}
    for layer in resolver.chain(sid):
        for region in layer.findall(Q('tblStylePr')):
            typ=region.get(Q('type'));dst=conditions.setdefault(typ,E.Element(Q('tblStylePr')))
            merge(dst,region)
    result['conditional']={k:canonical(v) for k,v in conditions.items()}
    # Preserve the template's next-paragraph editing convention when specified.
    nxt=style.find(Q('next'));result['next']=nxt.get(Q('val')) if nxt is not None else sid
    return result


def equivalent_styles(template,target,sid):
    a,_,_=package(template);b,_,_=package(target)
    return style_signature(a,sid) is not None and style_signature(a,sid)==style_signature(b,sid)


def equivalent_theme(template,target,part='word/theme/theme1.xml'):
    a,_,_=package(template);b,_,_=package(target)
    return part in a and part in b and canonical(parse(a[part]))==canonical(parse(b[part]))
