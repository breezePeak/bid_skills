#!/usr/bin/env python3
"""Conservatively clear only high-confidence bulk body direct-format drift."""
from __future__ import annotations
import argparse, json, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from template_style_profile import extract_profile

W="http://schemas.openxmlformats.org/wordprocessingml/2006/main"; NS={"w":W}
def qn(n): return f"{{{W}}}{n}"
def text(p): return ''.join(t.text or '' for t in p.findall('.//w:t',NS))
def sid(p):
    n=p.find('w:pPr/w:pStyle',NS); return n.get(qn('val')) if n is not None else None

def enabled(rpr,tag):
    if rpr is None:return False
    n=rpr.find(f'w:{tag}',NS)
    if n is None:return False
    return n.get(qn('val')) not in {'0','false','off'}

def repair(template:Path,src:Path,out:Path):
    profile=extract_profile(template); roles=profile.get('semantic_roles',{}); body_sid=(roles.get('body') or {}).get('style_id')
    heading_sids={v.get('style_id') for k,v in roles.items() if k.startswith('heading') and isinstance(v,dict) and v.get('style_id')}
    body_style=profile.get('styles',{}).get(body_sid,{}) if body_sid else {}; expected_size=(body_style.get('run') or {}).get('size_half_points')
    with zipfile.ZipFile(src) as z: root=ET.fromstring(z.read('word/document.xml'))
    table_paras={id(p) for tbl in root.findall('.//w:tbl',NS) for p in tbl.findall('.//w:p',NS)}
    changes=[]; in_body=False
    for idx,p in enumerate(root.findall('.//w:p',NS),1):
        psid=sid(p)
        if psid in heading_sids: in_body=True
        if not in_body or id(p) in table_paras or psid not in {None,body_sid}: continue
        raw=text(p).strip()
        if len(raw)<12: continue
        runs=[]; total=0; cov={'b':0,'i':0,'sz':0}
        for r in p.findall('w:r',NS):
            rt=''.join(t.text or '' for t in r.findall('.//w:t',NS)); n=len(rt)
            if not n: continue
            total+=n; rpr=r.find('w:rPr',NS); runs.append((r,rpr,n))
            if enabled(rpr,'b'): cov['b']+=n
            if enabled(rpr,'i'): cov['i']+=n
            if rpr is not None:
                sn=rpr.find('w:sz',NS); val=sn.get(qn('val')) if sn is not None else None
                if val and expected_size and val != str(expected_size): cov['sz']+=n
        if not total: continue
        props=[k for k,v in cov.items() if v/total>=0.60]
        if not props: continue
        removed=0
        for r,rpr,n in runs:
            if rpr is None: continue
            tags=[]
            if 'b' in props: tags+=['b','bCs']
            if 'i' in props: tags+=['i','iCs']
            if 'sz' in props: tags+=['sz','szCs']
            for tag in tags:
                node=rpr.find(f'w:{tag}',NS)
                if node is not None: rpr.remove(node); removed+=1
            if len(rpr)==0: r.remove(rpr)
        if removed: changes.append({'paragraph':idx,'properties':props,'removed_nodes':removed,'text':raw[:160]})
    xml=ET.tostring(root,encoding='utf-8',xml_declaration=True); out.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(src) as z, zipfile.ZipFile(out,'w') as dst:
        for info in z.infolist(): dst.writestr(info, xml if info.filename=='word/document.xml' else z.read(info.filename))
    return {'input':str(src),'output':str(out),'change_count':len(changes),'changes':changes}

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('template',type=Path); ap.add_argument('input',type=Path); ap.add_argument('--out',type=Path,required=True); ap.add_argument('--json-out',type=Path); a=ap.parse_args()
    r=repair(a.template,a.input,a.out); s=json.dumps(r,ensure_ascii=False,indent=2)
    if a.json_out: a.json_out.parent.mkdir(parents=True,exist_ok=True); a.json_out.write_text(s+'\n',encoding='utf-8')
    print(s); return 0
if __name__=='__main__': raise SystemExit(main())
