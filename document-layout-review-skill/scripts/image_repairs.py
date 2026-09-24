"""Authorize from actual defects before repainting; verify repaired pixels locally."""
from __future__ import annotations
import copy
from pathlib import Path
from block_progress import BlockError, sha, read_json, write_json


def _objects(path):
    from numbering_policy import Doc
    from review_objects import inventory
    doc=Doc(Path(path))
    return doc,{o['id']:o for o in inventory(doc) if o['kind']=='figure'}


def authorize(work,state,block):
    from visual_evidence import validate_inspection, INTRINSIC
    rows={r['id']:r for r in read_json(state['initial_review'])['objects']}
    _,current=_objects(state['current']);result=[]
    discovered=read_json(state['discovery_ledger']) if state.get('discovery_ledger') else {'inspections':[]}
    for ident in block.get('figures',[]):
        old=rows[ident];inspection=validate_inspection(old['inspection'],phase='initial')
        proofs=[]
        if inspection['verdict']=='fail' and old['object_sha256']==current[ident]['hash'] and any(inspection['checks'].get(k)=='fail' for k in INTRINSIC):
            proofs.append(old['inspection'])
        for ref in discovered.get('inspections',[]):
            found=validate_inspection(ref)
            if (found['binding'].get('object_id')==ident and found['binding'].get('object_sha256')==current[ident]['hash'] and found['verdict']=='fail'
                    and any(f['check'] in INTRINSIC for f in found['findings'])):
                proofs.append(ref)
        row={'id':ident,'current_object_sha256':current[ident]['hash'],'redraw_allowed':bool(proofs),'defect_evidence':proofs}
        if not proofs:row['instruction']='无已确认图内缺陷，保留原图；需要时只调整页面位置或等比大小。'
        result.append(row)
    plan={'version':1,'source_sha256':state['source_sha256'],'current_sha256':state['current_sha256'],
          'block':block['id'],'objects':result}
    dest=Path(work)/'image-plans'/(block['id']+'.json');write_json(dest,plan)
    return {'status':'image_plan_ready','plan':str(dest),'objects':result}


def _permission(work,state,block,ident,current_hash):
    from visual_evidence import validate_inspection, INTRINSIC
    path=Path(work)/'image-plans'/(block['id']+'.json')
    if not path.is_file():raise BlockError('先执行当前块 image-plan；没有缺陷依据时不得开始重画。')
    plan=read_json(path)
    if plan.get('source_sha256')!=state['source_sha256'] or plan.get('current_sha256')!=state['current_sha256']:
        raise BlockError('修图依据已过期，请重新核对当前块。')
    row=next((r for r in plan['objects'] if r['id']==ident),None)
    if not row or row.get('redraw_allowed') is not True or row.get('current_object_sha256')!=current_hash:
        raise BlockError('当前图片没有允许重绘的图内缺陷：'+ident)
    valid=False
    for ref in row.get('defect_evidence',[]):
        proof=validate_inspection(ref)
        if (proof['binding'].get('source_sha256')==state['source_sha256'] and proof['binding'].get('object_id')==ident
                and proof['binding'].get('object_sha256')==current_hash and proof['verdict']=='fail' and any(proof['checks'].get(k)=='fail' for k in INTRINSIC)):
            valid=True
    if not valid:raise BlockError('修图计划缺少真实失败依据。')


def inspect_repaired(work,state,block,proposal,config):
    import visual_evidence as v
    from figure_inspection import _extract, binding, _assert_pixels
    rows={r['id']:r for r in read_json(state['initial_review'])['objects']}
    current_doc,current=_objects(state['current']);doc,objects=_objects(proposal)
    completed=[]
    for ident in block.get('figures',[]):
        obj=objects.get(ident)
        if obj is None:raise BlockError('不能删除图片对象。')
        if obj['hash']==current[ident]['hash']:continue
        _permission(work,state,block,ident,current[ident]['hash'])
        root=Path(work)/'local-image-review'/ident
        old=v.validate_inspection(rows[ident]['inspection'],phase='initial')
        originals=v.validate_bundle(old['bundle'])['assets']
        local_row={}
        bound=root/'native-review.bound.json'
        if bound.is_file():
            local_row=next((r for r in v.read(bound).get('objects',[]) if r.get('id')==ident),{})
        if not obj.get('paths') and not local_row.get('rendered_view'):
            template=root/'native-review.json';root.mkdir(parents=True,exist_ok=True)
            v.write(template,{'objects':[{'id':ident,'object_sha256':obj['hash']}]})
            raise v.VisualError('visual-needs-rendered-crop','原生修复对象须绑定当前拟稿的渲染裁片；用已有 bind-rendered 输出 native-review.bound.json 后继续。',review=str(template),output=str(bound))
        assets=_extract(doc,obj,root,proposal,local_row)
        selected=[]
        if local_row.get('rendered_view'):
            from figure_inspection import _rendered_pixels
            _,selected=_rendered_pixels(proposal,local_row)
        task=v.prepare(assets,root,binding(state['source'],proposal,obj,doc,pages=selected,rendered_view=local_row.get('rendered_view')),phase='repair',pages=selected,
                       originals=[v.checked_file(a) for a in originals])
        failures=root/'failures.json'
        history=v.read(failures) if failures.is_file() else []
        now_pixels=[v.pixel_sha(v.normalized_image(v.checked_file(a))) for a in v.validate_bundle(task)['assets']]
        for failed_ref in history:
            failed=v.validate_inspection(failed_ref,phase='repair')
            old_pixels=[v.pixel_sha(v.normalized_image(v.checked_file(a))) for a in v.validate_bundle(failed['bundle'])['assets']]
            if failed['verdict']=='fail' and any(failed['checks'].get(k)=='fail' for k in v.INTRINSIC) and now_pixels==old_pixels:
                raise BlockError('修后图片已登记图内缺陷但像素未变；不能重复检查消除失败：'+ident)
        ref=v.run(task,config,root)
        proof=v.validate_inspection(ref,phase='repair')
        if proof['verdict']!='pass':
            history.append(ref);v.write(failures,history)
        dest=root/'latest.json';v.write(dest,{'inspection':ref,'rendered_view':local_row.get('rendered_view'),'pages':selected})
        if proof['verdict']!='pass':raise BlockError('修后的图片仍有缺陷或不确定，留在当前块：'+ident)
        completed.append(ident)
    return {'status':'local_images_checked','figures':completed,'page_layout':'pending-final'}


def require_repaired(work,state,block,proposal):
    from visual_evidence import validate_inspection, validate_bundle
    from figure_inspection import _assert_pixels, binding
    _,current=_objects(state['current']);doc,objects=_objects(proposal)
    for ident in block.get('figures',[]):
        obj=objects.get(ident)
        if obj is None:raise BlockError('不能删除图片对象。')
        if obj['hash']==current[ident]['hash']:continue
        _permission(work,state,block,ident,current[ident]['hash'])
        path=Path(work)/'local-image-review'/ident/'latest.json'
        if not path.is_file():raise BlockError('修图后先执行 inspect-repaired，不能只凭图片哈希变化就推进。')
        record=read_json(path)
        proof=validate_inspection(record['inspection'],phase='repair')
        wanted=binding(state['source'],proposal,obj,doc,pages=record.get('pages',[]),rendered_view=record.get('rendered_view'))
        comparable=lambda x:{k:v for k,v in x.items() if k!='candidate_sha256'}
        if proof['verdict']!='pass' or comparable(proof['binding'])!=comparable(wanted):
            raise BlockError('图片局部复查不属于当前真实图像。')
        _assert_pixels(doc,obj,validate_bundle(proof['bundle']),record.get('pages',[]),proposal,record)
