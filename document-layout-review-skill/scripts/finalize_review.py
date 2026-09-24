#!/usr/bin/env python3
"""Release only a hash-bound, field-updated, page- and image-reviewed candidate."""
from __future__ import annotations
import argparse
import json
import os
import shutil
import tempfile
from pathlib import Path
from numbering_policy import PolicyError, file_digest, digest
from template_caption_policy import audit_document
from figure_review import validate_final
from field_refresh import validate_report
from review_pipeline import REQUIRED_GATES
from render_docx import validate_render
from office_backends import OfficeError


def read(path):return json.loads(Path(path).read_text(encoding='utf-8-sig'))


def review_findings(value):
    result=[]
    def visit(x):
        if isinstance(x,dict):
            if x.get('severity')=='review':result.append(x)
            else:
                for v in x.values():visit(v)
        elif isinstance(x,list):
            for v in x:visit(v)
    visit(value)
    return {digest(json.dumps(x,ensure_ascii=False,sort_keys=True).encode()):x for x in result}


def validate_release(manifest, visual):
    m=read(manifest) if not isinstance(manifest,dict) else manifest
    v=read(visual) if not isinstance(visual,dict) else visual
    from block_progress import validate_final_binding
    validate_final_binding(m)
    if m.get('version')!=5 or m.get('status')!='awaiting_visual_review':
        raise PolicyError('manifest-not-ready','必须使用新版完整验收流水线，不能用中间状态或旧清单交付。')
    candidate=Path(m['candidate']);source=Path(m['source'])
    if not candidate.is_file() or file_digest(candidate)!=m.get('candidate_sha256'):
        raise PolicyError('candidate-stale','最后一次审计后 Word 又被修改。')
    if not source.is_file() or file_digest(source)!=m.get('source_sha256'):
        raise PolicyError('source-stale','原始 Word 已变，不能对照旧的初检记录。')
    if v.get('candidate_sha256')!=m['candidate_sha256'] or v.get('overall_status')!='pass':
        raise PolicyError('visual-not-passed','最终视觉检查未通过或不属于当前文件。')
    gates=m.get('gates')
    if not isinstance(gates,list) or len(gates)!=len(REQUIRED_GATES) or {g.get('id') for g in gates}!=REQUIRED_GATES:
        raise PolicyError('required-gate-missing','有必需检查被漏掉，不能交付。')
    gate_map={}
    for gate in gates:
        if gate.get('passed') is not True:raise PolicyError('gate-failed','仍有确定性检查失败。',gate=gate.get('id'))
        report=Path(gate['report'])
        if not report.is_file() or file_digest(report)!=gate.get('report_sha256'):
            raise PolicyError('gate-report-stale','审计报告缺失或被修改。',gate=gate['id'])
        gate_map[gate['id']]=read(report)
    validated_field=validate_report(candidate,gate_map['field-update'])
    required={'word':'Microsoft Word','wps':'WPS Writer','libreoffice':'LibreOffice UNO'}.get(m.get('field_engine_requested','auto'))
    if required and validated_field.get('engine')!=required:
        raise PolicyError('field-engine-mismatch','实际域更新引擎与明确要求不一致。',expected=required,actual=validated_field.get('engine'))
    for path_key,hash_key in [('template','template_sha256'),('template_style_json','template_style_sha256'),('initial_visual_review','initial_visual_review_sha256'),('content_plan','content_plan_sha256'),('text_rules','text_rules_file_sha256'),('runtime_preflight','runtime_preflight_sha256')]:
        path=Path(m[path_key])
        if not path.is_file() or file_digest(path)!=m.get(hash_key):raise PolicyError('baseline-stale','模板或初检基准发生变化。',path=str(path))
    render=m.get('render') or {};pages=render.get('pages') or []
    if render.get('passed') is not True or not pages or len(pages)!=len({p.get('name') for p in pages}):
        raise PolicyError('render-missing','必须渲染最新文件，且页面列表不能空或重复。')
    try:
        renderer_result=validate_render(candidate,render.get('data'),m.get('renderer_requested','auto'),pages)
    except OfficeError as exc:
        raise PolicyError(exc.code,str(exc),**exc.details) from exc
    expected={p['name']:p for p in pages};rows=v.get('pages')
    if not isinstance(rows,list) or len(rows)!=len(expected) or {p.get('name') for p in rows}!=set(expected):
        raise PolicyError('page-review-incomplete','最终页面没有逐页完整检查。')
    for row in rows:
        actual=expected[row['name']];path=Path(actual['path'])
        if not path.is_file() or file_digest(path)!=actual['sha256'] or row.get('sha256')!=actual['sha256']:
            raise PolicyError('page-review-stale','页面截图或检查记录已过期。',page=row['name'])
        if row.get('status')!='pass' or not str(row.get('observations','')).strip():
            raise PolicyError('page-review-not-passed','页面缺少实际观察或未通过。',page=row['name'])
    # Rerun the highest-risk policy on the current bytes, not merely a stored PASS.
    checked=audit_document(candidate,Path(m['template']),Path(m['template_style_json']),m.get('object_plan'))
    if checked['status']!='passed':raise PolicyError('final-numbering-failed','交付前重新核验题注/标题失败。',issues=checked['issues'])
    from stable_table_audit import audit as audit_table_template
    table_checked=audit_table_template(candidate,Path(m['template']),Path(m['template_style_json']))
    if table_checked['status']!='passed':raise PolicyError('final-table-template-failed','交付前表格外观不符合模板。',issues=table_checked['issues'])
    from content_integrity import audit as audit_content
    from template_format_contract import audit_extra_formats
    from template_layout_contract import audit as audit_table_layout
    rechecks = {
        'content-integrity': audit_content(source,candidate,m['content_plan']),
        'template-effective-format': audit_extra_formats(Path(m['template']),candidate,Path(m['template_style_json'])),
        'table-layout-template': audit_table_layout(Path(m['template']),candidate,Path(m['template_style_json'])),
    }
    for name,result in rechecks.items():
        if result.get('status')!='passed':
            raise PolicyError('final-'+name+'-failed','交付前当前文档重新核验未通过。',gate=name,issues=result.get('issues',[]))
    from visual_evidence import checked_file, VisualError
    try:
        if (v.get('figures') or {}).get('objects') or m.get('image_discovery_ledger'):
            checked_file(m.get('image_discovery_ledger'))
    except VisualError as exc:
        raise PolicyError(exc.code,str(exc),**exc.details) from exc
    image_result=validate_final(source,candidate,m['initial_visual_review'],v.get('figures') or {},pages,m.get('image_discovery_ledger'),v.get('pages'))
    from block_evidence import validate_table_coverage
    validate_table_coverage(candidate,v,pages)
    outstanding=review_findings(gate_map['table-layout']);decisions=v.get('table_reviews',[])
    if not isinstance(decisions,list) or len(decisions)!=len(outstanding) or {r.get('issue_sha256') for r in decisions}!=set(outstanding):
        raise PolicyError('table-review-missing','表格语义 review 候选未逐项说明是否需要合并或保留。',expected=list(outstanding))
    for row in decisions:
        if row.get('decision') not in {'keep_separate','acceptable'} or not str(row.get('reason','')).strip():
            raise PolicyError('table-review-unresolved','表格 review 未解决；需要修改的对象须先修复并重跑审计。')
    return {'status':'passed','candidate':str(candidate),'candidate_sha256':m['candidate_sha256'],
            'renderer':renderer_result,'page_count':len(pages),'figures':image_result,'heading_caption_recheck':checked,'table_template_recheck':table_checked,'final_rechecks':rechecks}


def finalize(manifest,visual,out):
    result=validate_release(manifest,visual);candidate=Path(result['candidate']);out=Path(out)
    m=read(manifest) if not isinstance(manifest,dict) else manifest
    protected={candidate.resolve(),Path(m['source']).resolve(),Path(m['template']).resolve()}
    for key in ('template_style_json','content_plan','text_rules','initial_visual_review','runtime_preflight'):
        if m.get(key):protected.add(Path(m[key]).resolve())
    if m.get('image_discovery_ledger'):protected.add(Path(m['image_discovery_ledger']['path']).resolve())
    for gate in m.get('gates',[]):protected.add(Path(gate['report']).resolve())
    for page in (m.get('render') or {}).get('pages',[]):protected.add(Path(page['path']).resolve())
    for key in ('pdf','report_path'):
        value=((m.get('render') or {}).get('data') or {}).get(key)
        if value:protected.add(Path(value).resolve())
    if not isinstance(manifest,dict):protected.add(Path(manifest).resolve())
    if not isinstance(visual,dict):protected.add(Path(visual).resolve())
    if out.resolve() in protected:raise PolicyError('output-collision','交付文件路径不能覆盖候选、原始 Word 或模板。')
    out.parent.mkdir(parents=True,exist_ok=True)
    with tempfile.NamedTemporaryFile(dir=out.parent,suffix='.docx',delete=False) as t:temp=Path(t.name)
    try:
        shutil.copyfile(candidate,temp)
        if file_digest(temp)!=result['candidate_sha256']:raise PolicyError('copy-mismatch','交付复制校验失败。')
        os.replace(temp,out)
    finally:temp.unlink(missing_ok=True)
    result.update(output=str(out),output_sha256=file_digest(out));return result


def main():
    ap=argparse.ArgumentParser();ap.add_argument('manifest',type=Path);ap.add_argument('visual_review',type=Path);ap.add_argument('--out',type=Path,required=True);ap.add_argument('--json-out',type=Path);a=ap.parse_args()
    if a.json_out and a.json_out.resolve() in {a.manifest.resolve(),a.visual_review.resolve(),a.out.resolve()}:
        print(json.dumps({'status':'failed','issues':[{'code':'report-output-collision','message':'报告不能覆盖清单、视觉记录或最终 Word。'}]},ensure_ascii=False));return 2
    try:result=finalize(a.manifest,a.visual_review,a.out)
    except Exception as exc:result={'status':'failed','output':None,'issues':[exc.as_issue() if hasattr(exc,'as_issue') else {'severity':'error','code':'release-gate-error','message':str(exc)}]}
    payload=json.dumps(result,ensure_ascii=False,indent=2)
    if a.json_out:a.json_out.parent.mkdir(parents=True,exist_ok=True);a.json_out.write_text(payload+'\n',encoding='utf-8')
    print(payload);return 0 if result['status']=='passed' else 2


if __name__=='__main__':raise SystemExit(main())

# DLR_BLOCK_WORKFLOW_V2
