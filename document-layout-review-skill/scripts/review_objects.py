"""Review inventory across body and auxiliary stories; numbering stays body-local."""
from __future__ import annotations
import copy
import posixpath
import re
from pathlib import Path


def inventory(doc):
    from numbering_policy import inventory as core_inventory
    rows=core_inventory(doc)
    for row in rows:row['part']='word/document.xml'
    for part in sorted(doc.files):
        if not re.fullmatch(r'word/(?:header[^/]*|footer[^/]*|footnotes|endnotes)\.xml',part):continue
        proxy=copy.copy(doc)
        proxy.files=dict(doc.files);proxy.roots=dict(doc.roots)
        proxy.document=doc.root(part)
        relpart=posixpath.join(posixpath.dirname(part),'_rels',posixpath.basename(part)+'.rels')
        mainrel='word/_rels/document.xml.rels'
        proxy.roots.pop(mainrel,None)
        if relpart in doc.files:
            proxy.files[mainrel]=doc.files[relpart]
        else:
            proxy.files.pop(mainrel,None)
        proxy.refresh()
        for row in core_inventory(proxy):
            row['id']='S_'+Path(part).stem+'_'+row['id']
            row['part']=part
            rows.append(row)
    return rows


def public_inventory(doc):
    from numbering_policy import visible
    return [{k:v for k,v in obj.items() if k not in {'element','caption'}} |
            {'caption_text':visible(obj['caption']) if obj['caption'] is not None else None,
             'caption_missing':obj['caption'] is None} for obj in inventory(doc)]
