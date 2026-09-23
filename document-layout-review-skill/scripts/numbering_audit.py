#!/usr/bin/env python3
"""Read-only gate for native heading, caption and footnote numbering."""
from __future__ import annotations
import argparse
import json
from pathlib import Path
from numbering_core import audit


def main():
    ap=argparse.ArgumentParser(description='检查标题、题注、脚注是否使用真实自动编号')
    ap.add_argument('input',type=Path);ap.add_argument('--json-out',type=Path)
    args=ap.parse_args()
    try:result=audit(args.input)
    except Exception as exc:result={'status':'failed','issues':[{'severity':'error','code':'numbering-audit-error','message':str(exc)}]}
    encoded=json.dumps(result,ensure_ascii=False,indent=2)
    if args.json_out:
        args.json_out.parent.mkdir(parents=True,exist_ok=True);args.json_out.write_text(encoded+'\n',encoding='utf-8')
    print(encoded)
    return 0 if result['status']=='passed' else 2


if __name__=='__main__':raise SystemExit(main())
