#!/usr/bin/env python3
"""Read-only whole-document inspection, including actual figure/caption inventory."""
from __future__ import annotations
import argparse
import json
import subprocess
import sys
from pathlib import Path
from numbering_policy import Doc, PolicyError, file_digest, public_inventory
from figure_review import make_review_template

HERE=Path(__file__).resolve().parent
BUNDLE=HERE.parent
DEFAULT_TEMPLATE=BUNDLE/'assets'/'default-template.docx'
DEFAULT_STYLE=BUNDLE/'assets'/'default-template-style.json'


def run_script(name,*args,allow=(0,2,3,4)):
    cp=subprocess.run([sys.executable,str(HERE/name),*map(str,args)],capture_output=True,text=True,check=False)
    return {'ok':cp.returncode in allow,'returncode':cp.returncode,'stdout':cp.stdout,'stderr':cp.stderr}


def load_json(path):
    try:return json.loads(Path(path).read_text(encoding='utf-8-sig'))
    except (OSError,ValueError):return None


def write(path,data):
    Path(path).parent.mkdir(parents=True,exist_ok=True)
    Path(path).write_text(json.dumps(data,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')


def main():
    ap=argparse.ArgumentParser(description='Word 第一轮只读检查，不把图片漏掉或把语义检查当图面通过')
    ap.add_argument('input',type=Path);ap.add_argument('--work-dir',type=Path,required=True)
    ap.add_argument('--template',type=Path);ap.add_argument('--template-style-json',type=Path)
    ap.add_argument('--text-rules',type=Path);args=ap.parse_args()
    reports=args.work_dir/'reports';reports.mkdir(parents=True,exist_ok=True)
    try:
        from review_pipeline import resolve_template
        from text_rules import load_rules,load_profile,expected_body_props,write_rules,rules_digest
        if not args.input.is_file():raise PolicyError('input-missing','输入不存在。')
        template,style_json=resolve_template(args,reports)
        rules=load_rules(args.text_rules);expected_body_props(load_profile(template),rules)
        rules_file=reports/'text-rules.effective.json';write_rules(rules_file,rules);rules_hash=rules_digest(rules)
        checks=[]
        calls=[
            ('结构与版式','docx_audit.py',[args.input],'docx-audit.json'),
            ('标题/题注/脚注自动编号','numbering_audit.py',[args.input,'--template',template,'--template-style-json',style_json],'automatic-numbering.json'),
            ('表格列宽/协调性/合并候选','table_layout_audit.py',[args.input,'--template-style-json',style_json],'table-layout.json'),
            ('中英文标点','contextual_punctuation.py',[args.input],'punctuation.json'),
            ('字符级文本','hard_text_audit.py',[args.input],'hard-text.json'),
            ('模板实际使用','template_usage_audit.py',[template,args.input,'--text-rules',rules_file],'template-usage.json'),
            ('模板结构一致性','template_conformance.py',[template,args.input],'template-conformance.json'),
        ]
        for name,script,params,filename in calls:
            path=reports/filename;rr=run_script(script,*params,'--json-out',path);data=load_json(path)
            if not rr['ok'] or data is None:
                raise PolicyError('inspection-execution-failed','初检程序执行失败，不能当作已检查。',script=script,run=rr)
            if script=='template_usage_audit.py' and data.get('text_rules_sha256')!=rules_hash:
                raise PolicyError('text-rules-not-used','文字检查未使用本次规则。')
            checks.append({'name':name,'report':str(path),'run':rr,'result':data})
        objects=public_inventory(Doc(args.input));inv=reports/'object-inventory.json'
        write(inv,{'version':1,'source_sha256':file_digest(args.input),'objects':objects})
        prepared=make_review_template(args.input,args.work_dir/'source-images')
        visual=reports/'initial-visual-review.template.json';write(visual,prepared)
        issues=[issue for c in checks for issue in (c['result'].get('issues') or [])]
        result={'status':'inspected','input':str(args.input),'input_sha256':file_digest(args.input),
                'template':str(template),'template_style_json':str(style_json),'text_rules':str(rules_file),'text_rules_sha256':rules_hash,
                'issue_count':sum(len(c['result'].get('issues',[])) if isinstance(c['result'].get('issues'),list) else c['result'].get('issue_count',0) for c in checks),
                'hard_error_count':sum(i.get('severity')=='error' for i in issues),'semantic_review_count':sum(i.get('severity')=='review' for i in issues),
                'checks':checks,'object_inventory':str(inv),'initial_visual_review_template':str(visual),
                'extracted_images':[a['path'] for o in prepared['objects'] for a in o['source_files']],
                'visual_review_required':True,
                'next':'Agent 查看全部真实原图；架构图和无箭头图也不能漏检。原生 Shape 没有媒体文件时渲染原 Word 检查。缺题注时依据图表内容填写 object-plan。检查清单里的 pending 不能当通过。'}
        write(args.work_dir/'inspection.json',result)
        print(json.dumps({'status':'inspected','inspection':str(args.work_dir/'inspection.json'),'object_inventory':str(inv),'initial_visual_review_template':str(visual),'figure_count':len(prepared['objects']),'hard_error_count':result['hard_error_count']},ensure_ascii=False,indent=2));return 0
    except Exception as exc:
        result={'status':'failed','issues':[exc.as_issue() if isinstance(exc,PolicyError) else {'severity':'error','code':'inspection-failed','message':str(exc)}]}
        write(args.work_dir/'inspection.json',result);print(json.dumps(result,ensure_ascii=False,indent=2));return 2


if __name__=='__main__':raise SystemExit(main())
