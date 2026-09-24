"""Reuse a real image inspection only if its image AND current pages are identical.

The original receipt is never rewritten. A new DOCX hash alone does not force a
new model call; changed pixels, page hashes, native-shape crops or failures do.
"""
from __future__ import annotations


def validate_bound_inspection(row, expected, *, phase):
    from visual_evidence import validate_inspection, VisualError
    result = validate_inspection(row.get('inspection'), phase=phase)
    bound = result['binding']
    if bound == expected:
        return result
    reuse = row.get('reuse') or {}
    comparable = lambda data: {k: v for k, v in data.items() if k != 'candidate_sha256'}
    if (phase != 'final' or expected.get('rendered_view') is not None or result['verdict'] != 'pass' or
            reuse.get('reason') != 'unchanged-image-and-page-hashes' or
            reuse.get('previous_candidate_sha256') != bound.get('candidate_sha256') or
            comparable(bound) != comparable(expected)):
        raise VisualError('visual-binding-mismatch', '旧图片结论不匹配当前对象/页面，须重新检查。')
    # verify_row subsequently compares actual current media and page pixels.
    return result


def reuse_source_content(source, candidate, obj, row, doc, original_row, pages, page_reviews, discoveries=None, *, validate=False):
    """Compose unchanged, really-inspected pixels with current full-page observations.

    This is NOT a fabricated final model call. It is explicitly labelled reuse;
    the release gate separately checks every current page. Any previous defect,
    changed image, native crop, or missing current-page observation rejects reuse.
    """
    from pathlib import Path
    from visual_evidence import (VisualError, checked_file, validate_inspection,
                                 validate_bundle, file_sha, read)
    from figure_inspection import verify_row, _assert_pixels
    if obj['hash'] != original_row.get('object_sha256') or original_row.get('rendered_view') or not obj.get('paths'):
        return False
    proof = validate_inspection(original_row.get('inspection'), phase='initial')
    if proof['verdict'] != 'pass' or proof['binding'].get('source_sha256') != file_sha(source):
        return False
    if discoveries:
        ledger = read(checked_file(discoveries))
        for receipt in ledger.get('inspections', []):
            found = validate_inspection(receipt, phase='final')
            if found['binding'].get('object_id') == obj['id']:
                return False  # existing defect closure always follows the normal repair verifier
    current_pages = {p['name']: p for p in pages}
    observations = {p['name']: p for p in (page_reviews or [])}
    refs = row.get('pages', [])
    if not refs or len({r.get('name') for r in refs}) != len(refs):
        return False
    actual = []
    for ref in refs:
        p = current_pages.get(ref.get('name')); review = observations.get(ref.get('name'), {})
        if (not p or ref.get('sha256') != p['sha256'] or review.get('sha256') != p['sha256']
                or review.get('status') != 'pass' or not str(review.get('observations', '')).strip()):
            return False
        checked_file(p)
        actual.append({'name': p['name'], 'sha256':p['sha256'], 'observations':review['observations']})
    bundle = validate_bundle(proof['bundle'])
    # It is legal to reuse only intrinsic pixels. Placement is explicitly checked above.
    _assert_pixels(doc, obj, bundle, [], candidate, {})
    expected = {'version':1, 'method':'unchanged-content-current-pages',
                'source_sha256':file_sha(source), 'candidate_sha256':file_sha(candidate),
                'object_sha256':obj['hash'], 'initial_inspection':original_row['inspection'], 'pages':actual}
    if validate:
        if row.get('content_review') != expected or row.get('checks') != proof['checks']:
            raise VisualError('content-review-stale','图内复用或当前页面记录不匹配。')
        return True
    row.update(content_review=expected, inspection=None, image_type=proof['image_type'],
               checks=proof['checks'], not_applicable_reasons=proof.get('not_applicable_reasons', {}),
               observation=proof['observation'], style_and_semantics_preserved=True,
               resolved_initial_defects=[], discovered_inspections=[], resolved_discovered_defects=[])
    row.pop('reuse', None)
    return True


def validate_table_coverage(candidate, visual, pages):
    """Every actual table is checked, not only rows flagged by a heuristic auditor."""
    from numbering_policy import Doc, inventory, PolicyError
    from review_objects import inventory
    expected = {o['id']:o for o in inventory(Doc(candidate)) if o['kind']=='table'}
    rows = visual.get('table_objects', [])
    if not isinstance(rows,list) or len(rows)!=len(expected) or {r.get('id') for r in rows}!=set(expected):
        raise PolicyError('table-object-review-incomplete','最终检查必须覆盖所有实际表格，而非只处理 review 候选。')
    actual_pages = {p['name']:p for p in pages}
    page_rows = {r['name']:r for r in visual.get('pages',[])}
    for row in rows:
        if row.get('object_sha256')!=expected[row['id']]['hash'] or row.get('status')!='pass' or not str(row.get('observations','')).strip():
            raise PolicyError('table-object-review-pending','表格缺少当前结果或实际观察。',object_id=row['id'])
        refs=row.get('pages',[])
        if not refs or len({r.get('name') for r in refs})!=len(refs):
            raise PolicyError('table-page-missing','表格必须定位到其所在最终页面。',object_id=row['id'])
        for ref in refs:
            p=actual_pages.get(ref.get('name')); review=page_rows.get(ref.get('name'),{})
            if not p or ref.get('sha256')!=p['sha256'] or review.get('sha256')!=p['sha256'] or review.get('status')!='pass':
                raise PolicyError('table-page-stale','表格所在页尚未通过当前页面检查。',object_id=row['id'])
    return {'status':'passed','table_count':len(expected)}
