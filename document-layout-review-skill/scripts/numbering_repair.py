#!/usr/bin/env python3
"""Convert identified manual numbers to native Word automatic numbering."""
from __future__ import annotations
import argparse
import json
from pathlib import Path
from numbering_core import repair


def main():
    ap=argparse.ArgumentParser(description='修复标题、题注和脚注自动编号；歧义不自动猜测')
    ap.add_argument('input',type=Path);ap.add_argument('--out',type=Path,required=True)
    ap.add_argument('--numbering-plan',type=Path,help='Agent 已确认的语义定位映射；须对应当前输入文件')
    ap.add_argument('--json-out',type=Path)
    args=ap.parse_args()
    try:result=repair(args.input,args.out,args.numbering_plan)
    except Exception as exc:result={'status':'blocked','input':str(args.input),'output':None,'change_count':0,'changes':[],
                                  'issues':[{'severity':'review','code':'numbering-needs-review','message':str(exc)}]}
    encoded=json.dumps(result,ensure_ascii=False,indent=2)
    if args.json_out:
        args.json_out.parent.mkdir(parents=True,exist_ok=True);args.json_out.write_text(encoded+'\n',encoding='utf-8')
    print(encoded)
    return 0 if result['status']=='passed' else 2


if __name__=='__main__':raise SystemExit(main())
