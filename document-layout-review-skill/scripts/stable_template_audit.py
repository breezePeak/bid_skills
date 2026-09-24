#!/usr/bin/env python3
"""Existing conformance rules with independently verified serialization equality."""
from __future__ import annotations
import argparse,json
from pathlib import Path
from block_progress import package_session
from format_equivalence import equivalent_styles,equivalent_theme


def check(template,target):
    from template_conformance import check as original_check
    result=original_check(Path(template),Path(target));kept=[];equivalent=[]
    with package_session():
        for issue in result['issues']:
            equal=False
            if issue.get('code')=='template-style-mutated':
                equal=equivalent_styles(template,target,issue['style_id'])
            elif issue.get('code')=='template-theme-changed':
                equal=equivalent_theme(template,target)
            (equivalent if equal else kept).append(issue)
    result.update(issues=kept,status='failed' if any(i.get('severity')=='error' for i in kept) else 'passed',
                  equivalent_serializations=equivalent)
    return result


def main():
    p=argparse.ArgumentParser();p.add_argument('template',type=Path);p.add_argument('target',type=Path);p.add_argument('--json-out',type=Path)
    a=p.parse_args()
    try:result=check(a.template,a.target)
    except Exception as exc:result={'status':'failed','issues':[{'code':'template-equivalence-error','message':str(exc)}]}
    text=json.dumps(result,ensure_ascii=False,indent=2)
    if a.json_out:
        if a.json_out.resolve() in {a.template.resolve(),a.target.resolve()}:raise SystemExit('报告不能覆盖文档')
        a.json_out.parent.mkdir(parents=True,exist_ok=True);a.json_out.write_text(text+'\n',encoding='utf-8')
    print(text);return 0 if result['status']=='passed' else 2


if __name__=='__main__':raise SystemExit(main())
