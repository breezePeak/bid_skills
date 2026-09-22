#!/usr/bin/env python3
"""Release gate: create final DOCX only after all structural and visual checks pass."""
from __future__ import annotations
import argparse, hashlib, json, shutil
from pathlib import Path

def sha256(path:Path):
    h=hashlib.sha256();
    with path.open('rb') as f:
        for c in iter(lambda:f.read(1024*1024),b''): h.update(c)
    return 'sha256:'+h.hexdigest()

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('manifest',type=Path); ap.add_argument('visual_review',type=Path); ap.add_argument('--out',type=Path,required=True); ap.add_argument('--json-out',type=Path); a=ap.parse_args()
    m=json.loads(a.manifest.read_text(encoding='utf-8')); v=json.loads(a.visual_review.read_text(encoding='utf-8'))
    errors=[]
    if m.get('status')!='awaiting_visual_review': errors.append('manifest is not awaiting_visual_review')
    if any(not g.get('passed') for g in m.get('gates',[])): errors.append('one or more structural gates failed')
    if not (m.get('render') or {}).get('passed'): errors.append('render gate failed')
    candidate=Path(m['candidate'])
    if not candidate.is_file() or sha256(candidate)!=m.get('candidate_sha256'): errors.append('candidate changed after audit')
    expected={p['name']:p['sha256'] for p in (m.get('render') or {}).get('pages',[])}
    reviewed={}
    for item in v.get('pages',[]): reviewed[item.get('name')]=item
    missing=sorted(set(expected)-set(reviewed)); extra=sorted(set(reviewed)-set(expected))
    if missing: errors.append('missing visual review pages: '+', '.join(missing))
    if extra: errors.append('unexpected visual review pages: '+', '.join(extra))
    for name,h in expected.items():
        item=reviewed.get(name)
        if item is None:
            continue
        if item.get('sha256')!=h: errors.append(f'page hash mismatch: {name}')
        if item.get('status')!='pass': errors.append(f'page did not pass visual review: {name}')
    if v.get('overall_status')!='pass': errors.append('overall visual review is not pass')
    result={'status':'failed' if errors else 'passed','errors':errors,'output':str(a.out) if not errors else None,'page_count':len(expected)}
    if not errors:
        a.out.parent.mkdir(parents=True,exist_ok=True); shutil.copy2(candidate,a.out)
        result['output_sha256']=sha256(a.out)
    payload=json.dumps(result,ensure_ascii=False,indent=2)
    if a.json_out: a.json_out.parent.mkdir(parents=True,exist_ok=True); a.json_out.write_text(payload+'\n',encoding='utf-8')
    print(payload); return 0 if not errors else 2
if __name__=='__main__': raise SystemExit(main())
