#!/usr/bin/env python3
"""Conservative useless-whitespace repair that preserves DOCX run structure.

Removes only high-confidence whitespace defects in Chinese prose. Characters are
removed from their original w:t nodes; runs, fonts, bold, size and paragraph
properties are not rebuilt.
"""
from __future__ import annotations
import argparse, json, re, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from punctuation_context import context_kind

W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'; NS={'w':W}
HAN=r'\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff'
CLOSE='，。；：！？、）】》」』”’'; OPEN='（【《「『“‘'

def text_nodes(p): return p.findall('.//w:t', NS)
def ptext(p): return ''.join(t.text or '' for t in text_nodes(p))

def removal_indexes(s: str) -> set[int]:
    if context_kind(s) not in {'chinese','mixed'} or not re.search(fr'[{HAN}]', s): return set()
    rm=set()
    # Leading/trailing spaces are formatting, not content, in ordinary Chinese prose.
    for m in re.finditer(r'^[ \u3000]+|[ \u3000]+$', s): rm.update(range(m.start(),m.end()))
    for m in re.finditer(fr'(?<=[{HAN}])[ \u3000]+(?=[{HAN}])', s): rm.update(range(m.start(),m.end()))
    for m in re.finditer(fr'[ \u3000]+(?=[{re.escape(CLOSE)}])', s): rm.update(range(m.start(),m.end()))
    for m in re.finditer(fr'(?<=[{re.escape(OPEN)}])[ \u3000]+', s): rm.update(range(m.start(),m.end()))
    return rm

def process_xml(data: bytes):
    root=ET.fromstring(data); changes=[]
    for idx,p in enumerate(root.findall('.//w:p',NS),1):
        nodes=text_nodes(p); before=''.join(t.text or '' for t in nodes)
        if not before: continue
        rm=removal_indexes(before)
        if not rm: continue
        offset=0; changed=0
        for t in nodes:
            old=t.text or ''; chars=[]
            for j,ch in enumerate(old):
                if offset+j not in rm: chars.append(ch)
            new=''.join(chars); offset += len(old)
            if new != old: t.text=new; changed += 1
        after=''.join(t.text or '' for t in nodes)
        changes.append({'paragraph':idx,'before':before,'after':after,'changed_text_nodes':changed})
    return ET.tostring(root,encoding='utf-8',xml_declaration=True), changes

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('input',type=Path); ap.add_argument('--out',type=Path,required=True); ap.add_argument('--json-out',type=Path); a=ap.parse_args()
    reps={}; changes=[]
    with zipfile.ZipFile(a.input) as z:
        parts=[n for n in z.namelist() if n=='word/document.xml' or (n.startswith('word/header') and n.endswith('.xml')) or (n.startswith('word/footer') and n.endswith('.xml'))]
        for part in parts:
            xml,ch=process_xml(z.read(part)); reps[part]=xml; changes += [{'part':part,**x} for x in ch]
        a.out.parent.mkdir(parents=True,exist_ok=True)
        with zipfile.ZipFile(a.out,'w') as dst:
            for info in z.infolist(): dst.writestr(info,reps.get(info.filename,z.read(info.filename)))
    result={'input':str(a.input),'output':str(a.out),'change_count':len(changes),'changes':changes}
    s=json.dumps(result,ensure_ascii=False,indent=2)
    if a.json_out: a.json_out.parent.mkdir(parents=True,exist_ok=True); a.json_out.write_text(s+'\n',encoding='utf-8')
    print(s); return 0
if __name__=='__main__': raise SystemExit(main())
