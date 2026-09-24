#!/usr/bin/env python3
"""Executable visual-review bridge for the existing document/object inventory.

The worker interface is provider-neutral. No model ids, object ids or prior PASS
results are guessed here. Existing numbering/content/style gates remain separate.
"""
from __future__ import annotations
import copy
import io
from pathlib import Path

from visual_evidence import (VisualError, CHECKS, INTRINSIC, checked_file, file_sha,
    reference, read, write, prepare, run, validate_inspection, validate_bundle,
    normalized_image, pixel_sha)


def _doc(path):
    from numbering_policy import Doc
    return Doc(path)


def _figures(doc):
    from numbering_policy import inventory
    from review_objects import inventory
    return [o for o in inventory(doc) if o['kind'] == 'figure']


def _media(doc, obj):
    from visual_evidence import sha
    return [sha(doc.files[name]) for name in obj['paths']]


def _page_refs(row, pages):
    by_name = {p['name']: p for p in pages}
    refs = row.get('pages')
    if not isinstance(refs, list) or not refs or not all(isinstance(r, dict) for r in refs) or len({r.get('name') for r in refs}) != len(refs):
        raise VisualError('visual-page-location-required', 'Agent 须先定位每张图所在的当前页面；不能跳过页面内效果检查。', object_id=row['id'])
    selected = []
    for ref in refs:
        p = by_name.get(ref.get('name'))
        if p is None or ref.get('sha256') != p['sha256']:
            raise VisualError('visual-page-stale', '图片定位引用了非当前页面。', object_id=row['id'])
        checked_file(p)
        selected.append(p)
    return sorted(selected, key=lambda p: p['name'])


def binding(source, current, obj, doc, *, pages=(), rendered_view=None):
    return {'source_sha256': file_sha(source), 'candidate_sha256': file_sha(current),
            'object_id': obj['id'], 'object_sha256': obj['hash'],
            'media_sha256': _media(doc, obj),
            'pages': [{'name': p['name'], 'sha256': p['sha256']} for p in pages],
            'rendered_view': rendered_view}


def _rendered_pixels(current, row):
    """Recreate native-shape/unsupported-media crops from a real bound render."""
    from render_docx import validate_render
    spec = row.get('rendered_view')
    if not isinstance(spec, dict):
        raise VisualError('visual-rendered-view-missing', '原生对象需要当前文档的已绑定渲染裁片。')
    report = read(checked_file(spec.get('render_report')))
    validate_render(current, report)
    pages = {Path(path).name: {**reference(path), 'name': Path(path).name} for path in report['pages']}
    locations = spec.get('locations')
    if not isinstance(locations, list) or not locations:
        raise VisualError('visual-crop-location', '渲染裁片必须指明实际页名和像素坐标。')
    images = []; selected = {}
    for loc in locations:
        page = pages.get(loc.get('name')) if isinstance(loc, dict) else None
        if page is None:
            raise VisualError('visual-crop-page', '裁片定位不属于当前渲染页面。')
        image = normalized_image(checked_file(page)); box = loc.get('bbox')
        if not (isinstance(box, list) and len(box) == 4 and all(type(x) is int for x in box)
                and 0 <= box[0] < box[2] <= image.width and 0 <= box[1] < box[3] <= image.height):
            raise VisualError('visual-crop-bounds', '对象裁片坐标必须在对应页面像素范围内。')
        images.append(image.crop(tuple(box))); selected[page['name']] = page
    return images, sorted(selected.values(), key=lambda p: p['name'])


def bind_rendered(current, review, render_report, locations):
    """Agent supplies only object ids/page names/bboxes; program generates hashes."""
    data = copy.deepcopy(review if isinstance(review, dict) else read(review))
    locations = locations if isinstance(locations, dict) else read(locations)
    rows = data.get('objects') if 'objects' in data else (data.get('figures') or {}).get('objects')
    if not isinstance(rows, list) or not isinstance(locations, dict):
        raise VisualError('visual-crop-map', '定位表应为对象 id 到页名/像素框列表的映射。')
    known = {o['id'] for o in _figures(_doc(current))}
    by_id = {r['id']: r for r in rows}
    if not set(locations).issubset(known & set(by_id)):
        raise VisualError('visual-crop-map', '不得猜测对象 id 或裁切不存在的对象。')
    for ident, locs in locations.items():
        row = by_id[ident]
        row['rendered_view'] = {'render_report': reference(render_report), 'locations': locs}
        _, pages = _rendered_pixels(current, row)
        if 'figures' in data:
            row['pages'] = [{'name': p['name'], 'sha256': p['sha256']} for p in pages]
    return data


def _assert_pixels(doc, obj, bundle, pages, current, row):
    if row.get('rendered_view') is not None:
        actual_images, located = _rendered_pixels(current, row)
        if not {p['name'] for p in located}.issubset({p['name'] for p in pages}):
            raise VisualError('visual-rendered-page-missing', '对象裁片所在页没有进入视觉检查。')
    else:
        if not obj['paths']:
            raise VisualError('visual-media-binding', '原生 Shape 须先绑定真实渲染裁片。', object_id=obj['id'])
        actual_images = [normalized_image(io.BytesIO(doc.files[name])) for name in obj['paths']]
    if len(actual_images) != len(bundle['assets']):
        raise VisualError('visual-media-binding', '检查视图没有完整对应文档内真实媒体或裁片。', object_id=obj['id'])
    for actual, asset in zip(actual_images, bundle['assets']):
        if pixel_sha(actual) != pixel_sha(normalized_image(checked_file(asset))):
            raise VisualError('visual-media-substituted', '检查图片不是当前 DOCX 内真实图像/裁片，不能用正常样例替代。', object_id=obj['id'])
    if len(bundle['pages']) != len(pages):
        raise VisualError('visual-page-coverage', '视觉调用没有检查对象所在的全部指定当前页。')
    for entry, actual in zip(bundle['pages'], pages):
        if entry['name'] != actual['name'] or entry['source_sha256'] != actual['sha256'] or pixel_sha(normalized_image(checked_file(entry))) != pixel_sha(normalized_image(checked_file(actual))):
            raise VisualError('visual-page-substituted', '视觉调用中的页面与当前渲染页不一致。')


def verify_row(source, current, obj, row, doc, *, phase, pages=(), original_row=None):
    selected = _page_refs(row, pages) if phase == 'final' else (_rendered_pixels(current, row)[1] if row.get('rendered_view') else [])
    expected = binding(source, current, obj, doc, pages=selected, rendered_view=row.get('rendered_view'))
    from block_evidence import validate_bound_inspection
    result = validate_bound_inspection(row, expected, phase=phase)
    bundle = validate_bundle(result['bundle'], result['binding'])
    _assert_pixels(doc, obj, bundle, selected, current, row)
    if phase == 'final':
        original_result = validate_inspection(original_row.get('inspection'), phase='initial')
        originals = validate_bundle(original_result['bundle'])['assets']
        if len(originals) != len(bundle['originals']):
            raise VisualError('visual-original-missing', '终检未与全部原图核对。')
        initial_intrinsic = {k for k, state in original_result['checks'].items() if state == 'fail'} & INTRINSIC
        current_pixels = [pixel_sha(normalized_image(checked_file(a))) for a in bundle['assets']]
        original_pixels = [pixel_sha(normalized_image(checked_file(a))) for a in originals]
        if initial_intrinsic and current_pixels == original_pixels:
            raise VisualError('image-pixels-unchanged', '已登记图内缺陷，但归一化像素未改变；重编码/改文件哈希不算修复。', object_id=obj['id'])
        for old, supplied in zip(originals, bundle['originals']):
            if pixel_sha(normalized_image(checked_file(old))) != pixel_sha(normalized_image(checked_file(supplied))):
                raise VisualError('visual-original-substituted', '终检使用的原图与原始 DOCX 初检不一致。')
    if row.get('checks') != result['checks'] or row.get('image_type') != result['image_type']:
        raise VisualError('visual-summary-mismatch', '图片检查表与真实视觉调用结论不一致。', object_id=obj['id'])
    if result['verdict'] == 'uncertain':
        raise VisualError('visual-uncertain', '有无法确认的视觉结果，须补充放大图/重新检查，不能记为通过。', object_id=obj['id'])
    if phase == 'final' and result['verdict'] != 'pass':
        raise VisualError('image-visual-defect', '终检新发现或未修复的图片问题必须返回修复。', object_id=obj['id'], findings=result['findings'])
    return result


def ledger(source, path):
    p = Path(path)
    if not p.is_file():
        write(p, {'version': 1, 'source_sha256': file_sha(source), 'inspections': []})
    data = read(p)
    if data.get('version') != 1 or data.get('source_sha256') != file_sha(source) or not isinstance(data.get('inspections'), list):
        raise VisualError('visual-discovery-stale', '新增缺陷台账不属于原始文档，不可重置来消除问题。')
    for ref in data['inspections']:
        outcome = validate_inspection(ref, phase='final')
        if outcome['binding']['source_sha256'] != data['source_sha256'] or outcome['verdict'] == 'pass':
            raise VisualError('visual-discovery-invalid', '新增缺陷台账包含不匹配或非失败记录。')
    return reference(p)


def _inspection_pixels(result):
    return [pixel_sha(normalized_image(checked_file(a))) for a in validate_bundle(result['bundle'])['assets']]


def late_findings(source, row, original_hash, current_hash, discovered, original_inspection=None):
    """Authorize a late-discovered repair without rewriting the initial baseline."""
    if discovered is None:
        refs = []
    else:
        data = read(checked_file(discovered))
        if data.get('version') != 1 or data.get('source_sha256') != file_sha(source):
            raise VisualError('visual-discovery-stale', '新增问题台账与原始文档不匹配。')
        refs = data.get('inspections')
        if not isinstance(refs, list):
            raise VisualError('visual-discovery-invalid', '新增问题台账格式不正确。')
    current_pixels = _inspection_pixels(validate_inspection(row['inspection'], phase='final')) if row.get('inspection') else None
    original_pixels = _inspection_pixels(validate_inspection(original_inspection, phase='initial')) if original_inspection else None
    own = []
    issues = []
    authorizes_change = False
    for ref in refs:
        res = validate_inspection(ref, phase='final')
        b = res['binding']
        if b['source_sha256'] != file_sha(source) or res['verdict'] == 'pass':
            raise VisualError('visual-discovery-invalid', '新增缺陷须来自同一原文的真实失败调用。')
        if b['object_id'] != row['id']:
            continue
        own.append(ref)
        prior_pixels = _inspection_pixels(res)
        if (b['object_sha256'] == original_hash or original_pixels is not None and prior_pixels == original_pixels) and res['findings']:
            authorizes_change = True
        for f in res['findings']:
            if f['check'] in INTRINSIC and (b['object_sha256'] == current_hash or current_pixels is not None and prior_pixels == current_pixels):
                raise VisualError('late-image-defect-unchanged', '终检已经发现图内缺陷，但真实图片未改变；不能靠新一轮 PASS 覆盖。', object_id=row['id'], issue=f['id'])
            issues.append(f['id'])
    if row.get('discovered_inspections', []) != own:
        raise VisualError('visual-discovery-omitted', '图片终检漏掉了已登记的新发现，不能只关闭初检问题。', object_id=row['id'])
    if set(row.get('resolved_discovered_defects', [])) != set(issues):
        raise VisualError('visual-discovery-unresolved', '后续发现的图片问题没有全部复查关闭。', object_id=row['id'])
    return {'issue_ids': issues, 'authorizes_change': authorizes_change}


def _extract(doc, obj, root, current=None, row=None):
    root = Path(root)
    root.mkdir(parents=True, exist_ok=True)
    files = []
    if row and row.get('rendered_view'):
        images, _ = _rendered_pixels(current, row)
        for index, image in enumerate(images, 1):
            p = root / f'rendered-object-{index}.png'
            image.save(p); files.append(p)
        return files
    for index, name in enumerate(obj['paths'], 1):
        p = root / f'media-{index}{Path(name).suffix}'
        p.write_bytes(doc.files[name])
        normalized_image(p)  # Do not let an unsupported picture silently disappear.
        files.append(p)
    if not files:
        raise VisualError('visual-needs-rendered-crop', '原生绘图没有媒体文件；先渲染源文档并提供对象所在页和裁切位置。', object_id=obj['id'])
    return files


def _apply_result(row, ref):
    result = validate_inspection(ref)
    row.update(inspection=ref, checks=result['checks'], image_type=result['image_type'],
               observation=result['observation'], not_applicable_reasons=result['not_applicable_reasons'])
    return result


def inspect_initial(source, review, config, out_dir, object_ids=None):
    data = copy.deepcopy(review if isinstance(review, dict) else read(review))
    doc = _doc(source)
    objects = _figures(doc)
    if data.get('source_sha256') != file_sha(source) or {r.get('id') for r in data.get('objects', [])} != {o['id'] for o in objects}:
        raise VisualError('initial-image-review-stale', '初检清单与原始 Word 不匹配。')
    rows = {r['id']: r for r in data['objects']}
    selected_ids = set(rows) if object_ids is None else set(object_ids)
    if not selected_ids <= set(rows):
        raise VisualError('unknown-image', '当前块包含未知图片 ID。')
    for obj in objects:
        if obj['id'] not in selected_ids:
            continue
        if rows[obj['id']].get('inspection'):
            try:
                verify_row(source, source, obj, rows[obj['id']], doc, phase='initial')
                continue
            except VisualError:
                pass
        root = Path(out_dir) / obj['id']
        row = rows[obj['id']]
        assets = _extract(doc, obj, root, source, row)
        selected = _rendered_pixels(source, row)[1] if row.get('rendered_view') else []
        task = prepare(assets, root, binding(source, source, obj, doc, pages=selected, rendered_view=row.get('rendered_view')), phase='initial', pages=selected)
        result = _apply_result(rows[obj['id']], run(task, config, root))
        if result['verdict'] == 'uncertain':
            data['status'] = 'requires_visual_review'
    data['version'] = 2
    data['status'] = 'inspected' if all(r.get('inspection') and all(v not in {'pending','uncertain'} for v in r.get('checks',{}).values()) for r in data['objects']) else 'requires_visual_review'
    return data


def inspect_final(source, candidate, initial, visual, manifest, config, out_dir):
    from figure_review import validate_initial
    old = validate_initial(source, initial)
    m = manifest if isinstance(manifest, dict) else read(manifest)
    if m.get('candidate_sha256') != file_sha(candidate) or m.get('source_sha256') != file_sha(source):
        raise VisualError('visual-manifest-stale', '最终视觉任务须使用本轮候选和清单。')
    from render_docx import validate_render
    validate_render(candidate, m['render']['data'], m.get('renderer_requested', 'auto'), m['render']['pages'])
    data = copy.deepcopy(visual if isinstance(visual, dict) else read(visual))
    if data.get('candidate_sha256') != file_sha(candidate):
        raise VisualError('visual-review-stale', '最终检查清单与本轮候选不匹配。')
    doc = _doc(candidate)
    objects = _figures(doc)
    rows = {r['id']: r for r in data.get('figures', {}).get('objects', [])}
    if len(rows) != len(data.get('figures', {}).get('objects', [])) or set(rows) != {o['id'] for o in objects} or set(old) != set(rows):
        raise VisualError('visual-object-coverage', '最终图片清单必须覆盖全部真实对象。')
    discovery_ref = m.get('image_discovery_ledger')
    if not discovery_ref:
        raise VisualError('visual-discovery-required', '请使用新版流水线生成新增缺陷台账，不能用旧清单放行。')
    discovery_path = checked_file(discovery_ref)
    discoveries = read(discovery_path)
    any_failure = False
    for obj in objects:
        row = rows[obj['id']]
        from block_evidence import reuse_source_content
        if reuse_source_content(source,candidate,obj,row,doc,old[obj['id']],m['render']['pages'],data.get('pages'),reference(discovery_path)):
            if not isinstance(visual,dict):
                write(visual,data)
            continue
        row.pop('content_review',None)
        if row.get('inspection'):
            try:
                verify_row(source, candidate, obj, row, doc, phase='final', pages=m['render']['pages'], original_row=old[obj['id']])
                late_findings(source, row, old[obj['id']]['object_sha256'], obj['hash'], reference(discovery_path), old[obj['id']]['inspection'])
                continue
            except VisualError:
                pass
        row.pop('reuse', None)
        root = Path(out_dir) / obj['id']
        assets = _extract(doc, obj, root, candidate, row)
        selected = _page_refs(row, m['render']['pages'])
        # A registered intrinsic defect in identical pixels requires repair,
        # not another paid model call hoping for a different verdict.
        current_pixels = [pixel_sha(normalized_image(p)) for p in assets]
        unchanged_failure = False
        for prior_ref in discoveries['inspections']:
            prior_result = validate_inspection(prior_ref, phase='final')
            if (prior_result['binding'].get('object_id') == obj['id'] and
                    prior_result['verdict'] == 'fail' and
                    any(prior_result['checks'].get(k) == 'fail' for k in INTRINSIC) and
                    _inspection_pixels(prior_result) == current_pixels):
                unchanged_failure = True
                break
        if unchanged_failure:
            any_failure = True
            row['postcheck_issue'] = {'code':'late-image-defect-unchanged',
                'message':'已登记图内缺陷但像素未变；不再次调用模型，先修当前图。'}
            row['resolved_discovered_defects'] = []
            if not isinstance(visual, dict):
                write(visual, data)
            continue
        original_result = validate_inspection(old[obj['id']]['inspection'], phase='initial')
        original_bundle = validate_bundle(original_result['bundle'])
        task = prepare(assets, root, binding(source, candidate, obj, doc, pages=selected, rendered_view=row.get('rendered_view')), phase='final',
                       pages=selected, originals=[checked_file(a) for a in original_bundle['assets']])
        proof = run(task, config, root)
        result = _apply_result(row, proof)
        row['style_and_semantics_preserved'] = result['semantics_preserved']
        if result['verdict'] != 'pass':
            any_failure = True
            discoveries['inspections'].append(proof)
        row['discovered_inspections'] = [ref for ref in discoveries['inspections'] if validate_inspection(ref)['binding']['object_id'] == obj['id']]
        row['resolved_initial_defects'] = [k for k, v in old[obj['id']]['checks'].items() if v == 'fail'] if result['verdict'] == 'pass' else []
        row['resolved_discovered_defects'] = [f['id'] for ref in row['discovered_inspections'] for f in validate_inspection(ref)['findings']] if result['verdict'] == 'pass' else []
        # A repair note is factual author input; never invent what the repair changed.
        row.setdefault('repair_note', '')
        write(discovery_path, discoveries)
        if result['verdict'] == 'pass':
            try:
                initial_bad = {k for k, v in old[obj['id']]['checks'].items() if v == 'fail'}
                if initial_bad & INTRINSIC and old[obj['id']]['object_sha256'] == obj['hash']:
                    raise VisualError('image-defect-unchanged', '初检图内缺陷已登记，但真实图片未改变。')
                late_findings(source, row, old[obj['id']]['object_sha256'], obj['hash'], reference(discovery_path), old[obj['id']]['inspection'])
            except VisualError as exc:
                any_failure = True
                row['postcheck_issue'] = exc.as_issue()
                row['resolved_initial_defects'] = []
                row['resolved_discovered_defects'] = []
        m['image_discovery_ledger'] = reference(discovery_path)
        # Save after each object so a later timeout cannot lose a newly found defect.
        if not isinstance(manifest, dict):
            write(manifest, m)
        if not isinstance(visual, dict):
            write(visual, data)
    data['figures']['version'] = 2
    data['figure_inspection_status'] = 'fail' if any_failure else 'pass'
    # This does NOT mark the whole document accepted; page/table checks remain mandatory.
    data['overall_status'] = 'fail' if any_failure else 'pending'
    return data

# DLR_BLOCK_WORKFLOW_V2
