#!/usr/bin/env python3
"""Repair -> real field update -> deterministic checks -> current-page review.

An audit-only continuation validates a manually repaired/refreshed candidate; it
never reruns broad repairs that could undo a reviewed diagram or pagination.
"""
from __future__ import annotations
import argparse
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path
from render_docx import validate_render
from numbering_policy import Doc, PolicyError, file_digest, inventory, public_inventory
from figure_review import make_review_template, validate_initial
from field_refresh import refresh, validate_report, safe_fields
from runtime_preflight import check as runtime_check, validate_docx, PreflightError
from content_integrity import make_plan, register_footnote_changes, ContentIntegrityError

HERE=Path(__file__).resolve().parent
BUNDLE=HERE.parent
DEFAULT_TEMPLATE=BUNDLE/'assets'/'default-template.docx'
DEFAULT_STYLE_JSON=BUNDLE/'assets'/'default-template-style.json'
REQUIRED_GATES={'numbering','caption-policy','punctuation','hard-text','structure',
                'template-text','template-structure','table-template','table-layout','image-inventory','field-update',
                'template-effective-format','table-layout-template','content-integrity'}


def write_json(path,data):
    path=Path(path);path.parent.mkdir(parents=True,exist_ok=True)
    path.write_text(json.dumps(data,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')


def load_json(path):
    try:return json.loads(Path(path).read_text(encoding='utf-8-sig'))
    except (OSError,ValueError,TypeError):return None


def _run_script_uncached(name,*args,allow=(0,)):
    cp=subprocess.run([sys.executable,str(HERE/name),*map(str,args)],capture_output=True,text=True,check=False,timeout=900 if name=='render_docx.py' else 300)
    return {'ok':cp.returncode in allow,'returncode':cp.returncode,'stdout':cp.stdout,'stderr':cp.stderr}


def run_script(name,*args,allow=(0,)):
    from block_cache import cached_run
    return cached_run(name,args,HERE,lambda n,*a:_run_script_uncached(n,*a,allow=allow))


def resolve_template(args,reports):
    if args.template is None:
        profile = args.template_style_json or DEFAULT_STYLE_JSON
        if not DEFAULT_TEMPLATE.is_file() or not profile.is_file():
            raise PolicyError('template-missing','默认模板或本次模板约定不存在。')
        return DEFAULT_TEMPLATE,profile
    if not args.template.is_file():raise PolicyError('template-missing','找不到本次模板。')
    if args.template_style_json is not None:
        if not args.template_style_json.is_file():raise PolicyError('template-profile-missing','找不到模板样式约定。')
        return args.template,args.template_style_json
    choice=reports/'template-choice-request.json'
    r=run_script('template_conflict_detector.py',args.template,'--json-out',choice,allow=(0,4))
    data=load_json(choice)
    if data and data.get('status')=='requires_user_choice':raise PolicyError('template-conflict','用户模板存在冲突，须先确定采用哪个约定。',choice_request=str(choice))
    if not r['ok']:raise PolicyError('template-inspection-failed','模板检查失败。',run=r)
    profile=reports/'uploaded-template.style-profile.json'
    r=run_script('template_style_profile.py',args.template,'--out',profile)
    if not r['ok'] or not profile.is_file():raise PolicyError('template-inspection-failed','无法解析用户模板样式。')
    return args.template,profile


def preserved_decisions(original,current,decisions):
    """Bind already-applied title/exemption decisions across deterministic stages.

    Figure identity and object counts must survive. Normal text cleanup may change
    a table text hash. Explicit paragraph locators are one-use and are not replayed.
    """
    old=inventory(Doc(original));new=inventory(Doc(current))
    if [o['kind'] for o in old]!=[o['kind'] for o in new]:raise PolicyError('object-inventory-changed','图表对象列表变了，不能复用旧语义决定。')
    for a,b in zip(old,new):
        if a['kind']=='figure' and a['hash']!=b['hash']:raise PolicyError('image-changed-by-batch-repair','文字/表格步骤擅自改变了图片。',object_id=a['id'])
    mapping={o['id']:o for o in new};rows=[]
    for row in decisions:
        if row.get('exempt'):
            rows.append({'id':row['id'],'object_sha256':mapping[row['id']]['hash'],'exempt':True,'reason':row['reason']})
    return {'version':1,'source_sha256':file_digest(current),'objects':rows}


def check_gates(candidate,template,style_json,rules_file,rules_hash,reports,object_plan,initial,source,field_report,content_plan):
    gates=[]
    def gate(ident,script,args,predicate):
        report=reports/('gate-'+ident+'.json')
        r=run_script(script,*args,'--json-out',report,allow=(0,2,3,4))
        data=load_json(report)
        ok=bool(r['returncode']==0 and data is not None and predicate(data))
        gates.append({'id':ident,'passed':ok,'report':str(report),'report_sha256':file_digest(report) if report.is_file() else None,'result':data,'run':r})
    extra=['--object-plan',object_plan] if object_plan else []
    gate('numbering','numbering_audit.py',[candidate,'--template',template,'--template-style-json',style_json,*extra],lambda d:d.get('status')=='passed')
    gate('caption-policy','template_caption_policy.py',['audit',candidate,'--template',template,'--template-style-json',style_json,*extra],lambda d:d.get('status')=='passed')
    gate('punctuation','contextual_punctuation.py',[candidate],lambda d:d.get('issue_count',0)==0)
    gate('hard-text','hard_text_audit.py',[candidate],lambda d:d.get('status')=='passed')
    gate('structure','docx_audit.py',[candidate],lambda d:not any(i.get('severity')=='error' for i in d.get('issues',[])))
    gate('template-text','template_usage_audit.py',[template,candidate,'--text-rules',rules_file],lambda d:d.get('status')=='passed' and d.get('text_rules_sha256')==rules_hash)
    gate('template-structure','template_conformance.py',[template,candidate],lambda d:d.get('status')=='passed')
    gate('table-template','template_table_style.py',['audit',candidate,'--template',template,'--template-style-json',style_json],lambda d:d.get('status')=='passed')
    gate('table-layout','table_layout_audit.py',[candidate,'--template-style-json',style_json],lambda d:not any(i.get('severity')=='error' for i in d.get('issues',[])))
    gate('template-effective-format','template_format_contract.py',[template,candidate,'--template-style-json',style_json],lambda d:d.get('status')=='passed')
    gate('table-layout-template','template_layout_contract.py',[template,candidate,'--template-style-json',style_json],lambda d:d.get('status')=='passed')
    gate('content-integrity','content_integrity.py',[source,candidate,'--content-plan',content_plan],lambda d:d.get('status')=='passed')
    rows=validate_initial(source,initial)
    actual=[o for o in inventory(Doc(candidate)) if o['kind']=='figure']
    data={'status':'passed' if {o['id'] for o in actual}==set(rows) else 'failed','objects':public_inventory(Doc(candidate)),'final_visual_review_required':bool(actual)}
    report=reports/'gate-image-inventory.json';write_json(report,data)
    gates.append({'id':'image-inventory','passed':data['status']=='passed','report':str(report),'report_sha256':file_digest(report),'result':data})
    try:data=validate_report(candidate,field_report);ok=True
    except PolicyError as exc:data={'status':'failed','issues':[exc.as_issue()]};ok=False
    report=reports/'gate-field-update.json';write_json(report,data)
    gates.append({'id':'field-update','passed':ok,'report':str(report),'report_sha256':file_digest(report),'result':data})
    return gates


def _audit_only_main():
    ap=argparse.ArgumentParser(description='按模板修复并执行不可缺项的图表/字段验收流程')
    ap.add_argument('input',type=Path);ap.add_argument('--work-dir',type=Path,required=True)
    ap.add_argument('--template',type=Path);ap.add_argument('--template-style-json',type=Path)
    ap.add_argument('--text-rules',type=Path);ap.add_argument('--numbering-plan',type=Path);ap.add_argument('--object-plan',type=Path)
    ap.add_argument('--initial-visual-review',type=Path)
    ap.add_argument('--vision-worker-config',type=Path,default=os.environ.get('DLR_VISION_WORKER_CONFIG'),help='已授权视觉模型/宿主子代理的命令适配器 JSON；不自动选择远程服务')
    ap.add_argument('--audit-only',action='store_true');ap.add_argument('--source',type=Path,help='audit-only 时必须指向原始 Word，而不是修改后的候选')
    ap.add_argument('--field-engine',choices=['auto','word','wps','libreoffice'],default='auto')
    ap.add_argument('--renderer',choices=['auto','word','wps','libreoffice'],default='auto')
    ap.add_argument('--field-update-report',type=Path,help='audit-only 时已更新域候选的引擎报告')
    ap.add_argument('--uno-python',help='可导入 uno 的 Python；不指定时自动探测')
    ap.add_argument('--content-plan',type=Path,help='继续验收时沿用原始文件绑定的内容保护计划')
    args=ap.parse_args()
    if not args.audit_only:
        ap.error('整本批量修复入口已停用；请通过分块入口推进，最后只审计。')
    for key in ('input','work_dir','template','template_style_json','text_rules','numbering_plan','object_plan','initial_visual_review','vision_worker_config','source','field_update_report','content_plan'):
        val=getattr(args,key)
        if val is not None:setattr(args,key,val.resolve())
    args.work_dir.mkdir(parents=True,exist_ok=True)
    reports=args.work_dir/'reports';reports.mkdir(exist_ok=True);render_dir=args.work_dir/'render'
    stages=[];source=args.source if args.audit_only else args.input
    try:
        if source is None or not source.is_file() or not args.input.is_file():raise PolicyError('input-missing','输入和原始 Word 必须存在。')
        reserved=['candidate.docx','candidate-refreshed.docx','00-numbering.docx','01-punctuation.docx','02-whitespace.docx','03-template-style.docx','04-template-usage.docx','04a-template-tables.docx','05-tables.docx','06-layout.docx','07-final-numbering.docx']
        if not args.audit_only and any((args.work_dir/name).resolve() in {args.input.resolve(),source.resolve()} for name in reserved):
            raise PolicyError('input-output-collision','工作目录中的阶段输出与输入重名；请选择新的工作目录，不能覆盖原文。')

        template,style_json=resolve_template(args,reports)
        if any((args.work_dir/name).resolve() in {template.resolve(),style_json.resolve()} for name in reserved):
            raise PolicyError('template-output-collision','阶段输出不能覆盖模板或模板约定。')
        preflight=runtime_check(source,template,args.work_dir,requested=args.field_engine,
                                profile=style_json,audit_only=args.audit_only,uno_python=args.uno_python,renderer=args.renderer)
        validate_docx(args.input)
        safe_fields(args.input)
        preflight_file=reports/'runtime-preflight.json';write_json(preflight_file,preflight)
        content_plan=reports/'content-plan.json'
        selected_plan=args.content_plan or (content_plan if args.audit_only and content_plan.is_file() else None)
        if selected_plan is not None:
            content_data=load_json(selected_plan)
            if not isinstance(content_data,dict) or content_data.get('version')!=1 or content_data.get('source_sha256')!=file_digest(source):
                raise PolicyError('content-plan-stale','内容保护计划不属于原始文件，不能把候选重新当作原文。')
        else:
            content_data=make_plan(source,args.object_plan)
        write_json(content_plan,content_data)

        # A hash-bound initial inventory forces image inspection before mutation.
        initial=args.initial_visual_review
        if initial is None:
            prepared=make_review_template(source,args.work_dir/'source-images')
            initial=reports/'initial-visual-review.json';write_json(initial,prepared)
            if prepared['objects']:
                write_json(reports/'object-inventory.json',{'version':1,'source_sha256':file_digest(source),'objects':public_inventory(Doc(source))})
                if args.vision_worker_config is not None:
                    from figure_inspection import inspect_initial
                    prepared=inspect_initial(source,prepared,args.vision_worker_config,reports/'initial-inspections')
                    write_json(initial,prepared)
                    validate_initial(source,initial)
                result={'status':'requires_image_review','review':str(initial),'inventory':str(reports/'object-inventory.json'),'message':'Agent 使用 figure_review.py inspect-initial 完成真实视觉调用并补齐缺题注语义计划后继续；不能手填 PASS。已传视觉配置时初检已自动调用，先读实际缺陷。'}
                write_json(args.work_dir/'review-manifest.json',result)
                print(json.dumps(result,ensure_ascii=False,indent=2));return 4
        validate_initial(source,initial)
        from text_rules import load_rules,load_profile,expected_body_props,write_rules,rules_digest
        rules=load_rules(args.text_rules);expected_body_props(load_profile(template),rules)
        rules_file=reports/'text-rules.effective.json';write_rules(rules_file,rules);rules_hash=rules_digest(rules)
        current=args.input;object_plan=args.object_plan
        def stage(ident,script,build):
            nonlocal current
            out=args.work_dir/(ident+'.docx');report=reports/(ident+'.json')
            if out.resolve()==current.resolve():raise PolicyError('stage-output-collision','阶段输出不能覆盖阶段输入。')
            from layout_invariant_guard import snapshot_docx,changed_invariants
            scope={'contextual_punctuation.py':'text-content','whitespace_repair.py':'text-content','template_style_enforce.py':'template-text-style','template_usage_repair.py':'template-text-style','docx_layout_policy.py':'table-layout'}.get(script)
            before=snapshot_docx(current) if scope else None
            r=run_script(script,*build(current,out,report));data=load_json(report)
            stages.append({'name':ident,'run':r,'report':str(report)})
            if not r['ok'] or not out.is_file():raise PolicyError('repair-stage-failed','确定性修复阶段失败。',stage=ident,result=data,run=r)
            if script in {'template_style_enforce.py','template_usage_repair.py'} and (data or {}).get('text_rules_sha256')!=rules_hash:
                raise PolicyError('text-rules-not-used','修复未使用本次文字规则。')
            if scope:
                damaged=changed_invariants(before,snapshot_docx(out),scope)
                stages[-1]['scope_guard']={'scope':scope,'passed':not damaged,'changed':damaged}
                if damaged:
                    out.unlink(missing_ok=True)
                    raise PolicyError('repair-scope-failed','本轮修复顺带改变了非目标对象；候选已撤销，须用更小范围修复。',stage=ident,invariants=damaged)
            current=out;return data or {}
        if not args.audit_only:
            numbering_args=['--numbering-plan',args.numbering_plan] if args.numbering_plan else []
            object_args=['--object-plan',object_plan] if object_plan else []
            first=stage('00-numbering','numbering_repair.py',lambda c,o,r:[c,'--out',o,'--json-out',r,'--template',template,'--template-style-json',style_json,*numbering_args,*object_args])
            numbered=current
            content_data=register_footnote_changes(source,numbered,first,content_data)
            write_json(content_plan,content_data)
            stage('01-punctuation','contextual_punctuation.py',lambda c,o,r:[c,'--out',o,'--json-out',r])
            stage('02-whitespace','whitespace_repair.py',lambda c,o,r:[c,'--out',o,'--json-out',r])
            stage('03-template-style','template_style_enforce.py',lambda c,o,r:[template,style_json,c,'--out',o,'--json-out',r,'--text-rules',rules_file])
            stage('04-template-usage','template_usage_repair.py',lambda c,o,r:[template,c,'--out',o,'--json-out',r,'--text-rules',rules_file])
            # Normalize from the TEMPLATE first; do not freeze erroneous source fills.
            stage('04a-template-tables','template_table_style.py',lambda c,o,r:['repair',c,'--out',o,'--json-out',r,'--template',template,'--template-style-json',style_json])
            stage('05-tables','table_layout_repair.py',lambda c,o,r:[c,'--output',o,'--json-out',r,'--template',template])
            stage('06-layout','docx_layout_policy.py',lambda c,o,r:[c,'--output',o,'--json-out',r,'--template',template,'--template-style-json',style_json])
            rebound=preserved_decisions(numbered,current,first.get('object_decisions',[]))
            object_plan=reports/'object-decisions.effective.json';write_json(object_plan,rebound)
            # The legacy template step must not reintroduce a source-derived list
            # or undo caption placement. Reapply the SAME template policy last.
            stage('07-final-numbering','numbering_repair.py',lambda c,o,r:[c,'--out',o,'--json-out',r,'--template',template,'--template-style-json',style_json,'--object-plan',object_plan])
        candidate=args.work_dir/'candidate.docx'
        if current.resolve()!=candidate.resolve():shutil.copy2(current,candidate)
        # In audit-only mode candidate bytes must already be the refreshed bytes.
        if args.audit_only:
            if args.field_update_report is None:raise PolicyError('field-update-report-missing','继续验收前必须实际更新最后修改后的 Word。')
            field_data=validate_report(candidate,args.field_update_report)
            required={'word':'Microsoft Word','wps':'WPS Writer','libreoffice':'LibreOffice UNO'}.get(args.field_engine)
            if required and field_data.get('engine')!=required:
                raise PolicyError('field-engine-mismatch','续跑时的真实域更新引擎不符合明确要求。',expected=required,actual=field_data.get('engine'))
        else:
            refreshed=args.work_dir/'candidate-refreshed.docx'
            try:
                field_data=refresh(candidate,refreshed,args.field_engine,args.uno_python,preflight=preflight['field_engine'])
                shutil.copy2(refreshed,candidate)
            except PolicyError as exc:
                result={'status':'requires_field_update','candidate':str(candidate),'candidate_sha256':file_digest(candidate),'issues':[exc.as_issue()],
                        'next':'修复实际引擎错误后执行 field_refresh.py（默认 Word→WPS→LibreOffice，明确指定时不降级）；随后 --audit-only 检查，不能将此候选当已验收成果。'}
                write_json(args.work_dir/'review-manifest.json',result)
                print(json.dumps(result,ensure_ascii=False,indent=2));return 4
        field_report=reports/'field-update.json';write_json(field_report,field_data)
        gates=check_gates(candidate,template,style_json,rules_file,rules_hash,reports,object_plan,initial,source,field_report,content_plan)
        render_run=run_script('render_docx.py',candidate,'--out-dir',render_dir,'--renderer',args.renderer)
        try:render_data=json.loads(render_run['stdout'])
        except ValueError:render_data={}
        pages=[{'path':str(Path(p).resolve()),'name':Path(p).name,'sha256':file_digest(Path(p))} for p in render_data.get('pages',[]) if Path(p).is_file()]
        render_valid = False
        if render_run['ok'] and pages:
            try:
                validate_render(candidate,render_data,args.renderer,pages)
                render_valid = True
            except Exception as exc:
                render_data['validation_error'] = exc.as_issue() if hasattr(exc,'as_issue') else {'message':str(exc)}
        ok=all(g['passed'] for g in gates) and render_valid
        from figure_inspection import ledger
        discovery_ref=ledger(source,reports/'image-discoveries.json')
        manifest={'image_discovery_ledger':discovery_ref,'version':5,'renderer_requested':args.renderer,'field_engine_requested':args.field_engine,'status':'awaiting_visual_review' if ok else 'failed','source':str(source.resolve()),'source_sha256':file_digest(source),
                  'candidate':str(candidate.resolve()),'candidate_sha256':file_digest(candidate),'template':str(template.resolve()),'template_sha256':file_digest(template),
                  'template_style_json':str(style_json.resolve()),'template_style_sha256':file_digest(style_json),
                  'text_rules':str(rules_file.resolve()),'text_rules_sha256':rules_hash,'text_rules_file_sha256':file_digest(rules_file),
                  'content_plan':str(content_plan.resolve()),'content_plan_sha256':file_digest(content_plan),
                  'runtime_preflight':str(preflight_file.resolve()),'runtime_preflight_sha256':file_digest(preflight_file),'object_plan':str(object_plan.resolve()) if object_plan else None,
                  'initial_visual_review':str(initial.resolve()),'initial_visual_review_sha256':file_digest(initial),'stages':stages,'gates':gates,
                  'render':{'passed':render_valid,'run':render_run,'data':render_data,'pages':pages},
                  'next':'Agent 定位每张图的当前页面，再用 figure_review.py inspect-final 进行两次独立完整盲检；先前通过不豁免新发现。缺陷台账须闭合，逐页与表格检查仍必需。最后使用 finalize_review.py。'}
        write_json(args.work_dir/'review-manifest.json',manifest)
        final_template={'overall_status':'pending','candidate_sha256':file_digest(candidate),'pages':[{'name':p['name'],'sha256':p['sha256'],'status':'pending','observations':''} for p in pages],
                        'figures':{'version':2,'source_sha256':file_digest(source),'candidate_sha256':file_digest(candidate),'objects':[]},'table_reviews':[]}
        initial_rows=validate_initial(source,initial)
        for obj in inventory(Doc(candidate)):
            if obj['kind']!='figure':continue
            old=initial_rows.get(obj['id'],{})
            final_template['figures']['objects'].append({'id':obj['id'],'object_sha256':obj['hash'],'source_object_sha256':old.get('object_sha256'),
                'image_type':old.get('image_type'),'observation':'','checks':{key:'pending' for key in old.get('checks',{})},
                'inspection':None,'discovered_inspections':[],'resolved_discovered_defects':[],
                'style_and_semantics_preserved':False,'resolved_initial_defects':[],'repair_note':'','pages':[]})
        # Issue IDs are generated, not invented by the model.
        table_result=next(g.get('result') or {} for g in gates if g['id']=='table-layout')
        from numbering_policy import digest
        for issue in table_result.get('issues',[]):
            if issue.get('severity')=='review':
                final_template['table_reviews'].append({'issue_sha256':digest(json.dumps(issue,ensure_ascii=False,sort_keys=True).encode()),'decision':'pending','reason':''})
        write_json(reports/'final-visual-review.template.json',final_template)
        print(json.dumps({'status':manifest['status'],'candidate':str(candidate),'manifest':str(args.work_dir/'review-manifest.json'),'visual_review_template':str(reports/'final-visual-review.template.json'),'failed_gates':[g['id'] for g in gates if not g['passed']]},ensure_ascii=False,indent=2))
        return 0 if ok else 2
    except Exception as exc:
        result={'status':'failed','issues':[exc.as_issue() if hasattr(exc,'as_issue') else {'severity':'error','code':'review-pipeline-error','message':str(exc)}],'stages':stages}
        write_json(args.work_dir/'review-manifest.json',result)
        write_json(reports/'pipeline-failure.json',result);print(json.dumps(result,ensure_ascii=False,indent=2));return 2


def main():
    from block_workflow import main as block_main
    return block_main()


if __name__=='__main__':raise SystemExit(main())

# DLR_BLOCK_WORKFLOW_V2
