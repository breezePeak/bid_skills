#!/usr/bin/env python3
"""Audit actual template style usage without penalizing equivalent direct formatting."""
from __future__ import annotations
import argparse, json, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from template_style_profile import extract_profile

W="http://schemas.openxmlformats.org/wordprocessingml/2006/main"; NS={"w":W}
def qn(n): return f"{{{W}}}{n}"
def p_text(p): return ''.join(t.text or '' for t in p.findall('.//w:t',NS))
def style_id(p):
    x=p.find('w:pPr/w:pStyle',NS); return x.get(qn('val')) if x is not None else None

def boolprop(rpr,tag):
    if rpr is None:return False
    n=rpr.find(f'w:{tag}',NS)
    if n is None:return False
    v=n.get(qn('val')); return v not in {'0','false','off'}

def audit(template:Path,target:Path):
    profile=extract_profile(template)
    roles=profile.get('semantic_roles',{})
    body_sid=(roles.get('body') or {}).get('style_id')
    heading_sids={v.get('style_id') for k,v in roles.items() if k.startswith('heading') and isinstance(v,dict) and v.get('style_id')}
    body_style=profile.get('styles',{}).get(body_sid,{}) if body_sid else {}
    expected_size=(body_style.get('run') or {}).get('size_half_points')
    with zipfile.ZipFile(target) as z: root=ET.fromstring(z.read('word/document.xml'))
    table_paras={id(p) for tbl in root.findall('.//w:tbl',NS) for p in tbl.findall('.//w:p',NS)}
    issues=[]; in_body=False
    for idx,p in enumerate(root.findall('.//w:p',NS),1):
        sid=style_id(p)
        if sid in heading_sids: in_body=True
        if not in_body or id(p) in table_paras: continue
        text=p_text(p).strip()
        if len(text)<12 or sid not in {None,body_sid}: continue
        total=0; bold=italic=size_diff=0; size_vals={}
        for r in p.findall('w:r',NS):
            rt=''.join(t.text or '' for t in r.findall('.//w:t',NS)); n=len(rt)
            if not n: continue
            total+=n; rpr=r.find('w:rPr',NS)
            if boolprop(rpr,'b'): bold+=n
            if boolprop(rpr,'i'): italic+=n
            if rpr is not None:
                sz=rpr.find('w:sz',NS)
                val=sz.get(qn('val')) if sz is not None else None
                if val and expected_size and val != str(expected_size):
                    size_diff+=n; size_vals[val]=size_vals.get(val,0)+n
        if not total: continue
        def add(code,covered,msg,evidence=None):
            if covered/total>=0.60:
                issues.append({'severity':'error','code':code,'paragraph':idx,'coverage':round(covered/total,3),'text':text[:160],'message':msg,'evidence':evidence})
        add('body-bulk-bold-direct-format',bold,'普通正文大面积使用直接加粗，覆盖了模板正文样式。')
        add('body-bulk-italic-direct-format',italic,'普通正文大面积使用直接斜体，覆盖了模板正文样式。')
        add('body-bulk-size-override',size_diff,'普通正文大面积使用与模板不同的直接字号。',size_vals)
    return {'template':str(template),'target':str(target),'status':'passed' if not issues else 'failed','issues':issues}

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('template',type=Path); ap.add_argument('target',type=Path); ap.add_argument('--json-out',type=Path); a=ap.parse_args()
    r=audit(a.template,a.target); s=json.dumps(r,ensure_ascii=False,indent=2)
    if a.json_out: a.json_out.parent.mkdir(parents=True,exist_ok=True); a.json_out.write_text(s+'\n',encoding='utf-8')
    print(s); return 0 if r['status']=='passed' else 2
if __name__=='__main__': raise SystemExit(main())
