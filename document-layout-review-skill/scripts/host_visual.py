"""Two-step host vision transport, using the same pixel and verdict contract.

The host opens the exported images, supplies its real tool/transcript references,
then resumes. Local receipts verify coverage/freshness, not unforgeable host
attestation. A task or checklist alone is never a completed visual inspection.
"""
from __future__ import annotations
import argparse
import copy
import json
import uuid
from pathlib import Path


def validate_log(log, request):
    from visual_evidence import VisualError
    if not isinstance(log,dict) or log.get('request_id')!=request['request_id']:
        raise VisualError('host-vision-log-missing','缺少与当前图片任务绑定的宿主看图记录。')
    calls=log.get('tool_calls')
    if not isinstance(calls,list) or not calls:
        raise VisualError('host-vision-log-missing','必须先实际调用宿主看图工具，不能只填写通过。')
    wanted={r['id']:r['sha256'] for r in request['images']};seen=set()
    for call in calls:
        if not isinstance(call,dict) or not str(call.get('tool','')).strip() or not str(call.get('reference','')).strip():
            raise VisualError('host-vision-log-missing','看图记录须包含实际工具名称和宿主日志/会话引用；不编造调用编号。')
        for row in call.get('images',[]):
            if not isinstance(row,dict) or wanted.get(row.get('id'))!=row.get('sha256'):
                raise VisualError('host-vision-image-mismatch','宿主看图记录引用了其他图片或过期像素。')
            seen.add(row['id'])
    if seen!=set(wanted):
        raise VisualError('host-vision-incomplete','宿主看图漏掉完整图、细节、原图或当前页面。',missing=sorted(set(wanted)-seen))


def run_host(bundle_ref, out_dir):
    import visual_evidence as v
    bundle=v.validate_bundle(bundle_ref)
    key=v.sha(v.canonical({'phase':bundle['phase'],'binding':bundle['binding'],
        'images':[(r['id'],r['sha256']) for r in bundle['views']+bundle['pages']+bundle['originals']]})).split(':')[1]
    root=Path(out_dir)/('host-'+key);root.mkdir(parents=True,exist_ok=True)
    report=root/'inspection.json';task=root/'task.json'
    if report.is_file():
        ref=v.reference(report)
        v.validate_inspection(ref,bundle['binding'],phase=bundle['phase'])
        return ref
    if not task.is_file():
        role=v.ROLES[bundle['phase']][0]
        request=v.make_request(bundle,role,uuid.uuid4().hex)
        request_path=root/'request.json';v.write(request_path,request)
        v.write(task,{'version':1,'status':'awaiting_host_visual','bundle':bundle_ref,'request':v.reference(request_path),
            'role':role,'request_id':request['request_id'],
            'images':[{'id':r['id'],'path':r['path'],'sha256':r['sha256']} for r in bundle['views']+bundle['pages']+bundle['originals']],
            'instruction':'使用宿主看图工具查看 images 的实际图片；按 request 的检查项目记录观察与缺陷。保存真实工具/会话引用，不使用 OCR 文字替代图像。随后运行 host_visual.py task.json response.json tool-log.json，再重试原分块命令。'})
    data=v.read(task);v.checked_file(data['bundle']);v.checked_file(data['request'])
    raise v.VisualError('visual-host-review-required','没有可调用的自动视觉接口时，转宿主看图；当前尚未检查，不是通过。',task=str(task),images=data['images'])


def complete_host(task, response, log, *, allow_test_double=False):
    import visual_evidence as v
    task=Path(task);data=v.read(task)
    if data.get('version')!=1 or data.get('status')!='awaiting_host_visual':
        raise v.VisualError('host-vision-task-invalid','不是有效的待检查宿主任务。')
    request=v.read(v.checked_file(data['request']))
    bundle=v.validate_bundle(data['bundle'])
    if request!=v.make_request(bundle,data['role'],data['request_id']):
        raise v.VisualError('host-vision-task-stale','宿主任务已改变，不能导入旧结论。')
    response=v.read(response) if isinstance(response,(str,Path)) else response
    log=v.read(log) if isinstance(log,(str,Path)) else log
    validate_log(log,request)
    if (response.get('_test_double') or log.get('_test_double')) and not allow_test_double:
        raise v.VisualError('visual-test-double-forbidden','模拟宿主记录不能用于正式放行。')
    result=v.evaluate(response,request)
    root=task.parent;report=root/'inspection.json'
    if report.exists():
        raise v.VisualError('host-vision-already-completed','原始看图结果不得覆盖；有新像素时生成新任务。')
    response_path=root/'host-response.json';log_path=root/'host-tools.json';stderr=root/'host-stderr.txt'
    v.write(response_path,response);v.write(log_path,log);stderr.write_text('',encoding='utf-8')
    call={'role':data['role'],'request_id':data['request_id'],'returncode':0,'transport':'host',
          'request':data['request'],'response':v.reference(response_path),'stderr':v.reference(stderr),
          'host_log':v.reference(log_path),'result':result}
    v.write(report,{'protocol':v.PROTOCOL,'status':'completed','bundle':data['bundle'],
        'phase':bundle['phase'],'binding':bundle['binding'],'calls':[call]})
    v.validate_inspection(v.reference(report),allow_test_double=allow_test_double)
    return {'status':'completed','verdict':result['verdict'],'inspection':v.reference(report)}


def main():
    p=argparse.ArgumentParser(description='导入宿主实际看图结论；待检查不算通过')
    p.add_argument('task',type=Path);p.add_argument('response',type=Path);p.add_argument('tool_log',type=Path)
    a=p.parse_args()
    try:
        result=complete_host(a.task,a.response,a.tool_log)
    except Exception as exc:
        result={'status':'blocked','issue':exc.as_issue() if hasattr(exc,'as_issue') else {'message':str(exc)}}
    print(json.dumps(result,ensure_ascii=False,indent=2))
    return 0 if result.get('verdict')=='pass' else 2


if __name__=='__main__':raise SystemExit(main())
