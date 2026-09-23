#!/usr/bin/env python3
"""Hash-bound visual review gate for real images, not a pixel/OCR classifier.

The Agent must inspect the complete source image and the latest rendered page.
This module verifies coverage, evidence freshness and issue closure. It never
infers that a bitmap's labels fit merely because its semantic graph is valid.
"""
from __future__ import annotations
import argparse
import json
from pathlib import Path
from numbering_policy import Doc, PolicyError, digest, file_digest, inventory

CHECKS = ('text_inside_bounds', 'no_overlap', 'readable', 'not_clipped',
          'connections_correct', 'no_embedded_caption')
INTRINSIC = {'text_inside_bounds', 'no_overlap', 'connections_correct', 'no_embedded_caption'}
IMAGE_TYPES = {'diagram', 'photo', 'decoration'}


def load(value):
    return value if isinstance(value, dict) else json.loads(Path(value).read_text(encoding='utf-8-sig'))


def figures(path):
    return [o for o in inventory(Doc(path)) if o['kind'] == 'figure']


def make_review_template(source, out_dir):
    source, out_dir = Path(source), Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    doc = Doc(source); items = []
    for obj in inventory(doc):
        if obj['kind'] != 'figure': continue
        extracted = []
        for i, name in enumerate(obj['paths'], 1):
            dest = out_dir / (obj['id'] + '-' + str(i) + Path(name).suffix)
            dest.write_bytes(doc.files[name])
            extracted.append({'path': str(dest), 'sha256': file_digest(dest)})
        items.append({'id': obj['id'], 'object_sha256': obj['hash'],
                      'image_type': None, 'observation': '',
                      'checks': {k: 'pending' for k in CHECKS},
                      'source_files': extracted, 'context': obj['context'],
                      'caption_outside_image': obj['caption'] is not None})
    return {'version': 1, 'source_sha256': file_digest(source), 'objects': items,
            'instruction': '逐张看完整原图；放大细看文字边界，再查看最终页。pending 不得改成未核验的 pass。架构图无需编造流程箭头。原图下方的 Word 图题不是图内题注。'}


def indexed(rows, actual):
    if not isinstance(rows, list): raise PolicyError('image-review-invalid', '图片检查对象必须为数组。')
    result = {}
    for row in rows:
        if not isinstance(row, dict) or row.get('id') not in actual or row['id'] in result:
            raise PolicyError('image-review-coverage', '图片检查含重复、未知或无效对象。')
        result[row['id']] = row
    if set(result) != set(actual):
        raise PolicyError('image-review-missing', '不能漏检图片。', missing=sorted(set(actual) - set(result)))
    return result


def validate_checks(row, *, final=False):
    typ = row.get('image_type')
    if typ not in IMAGE_TYPES: raise PolicyError('image-review-unclassified', '必须确认图片是图解、照片还是装饰图。', object_id=row['id'])
    if not str(row.get('observation', '')).strip():
        raise PolicyError('image-review-no-observation', '图片检查必须描述实际观察到的内容与边界。', object_id=row['id'])
    checks = row.get('checks')
    if not isinstance(checks, dict) or set(checks) != set(CHECKS):
        raise PolicyError('image-review-incomplete', '图片检查项目不完整。', object_id=row['id'])
    for name, state in checks.items():
        allowed_na = name == 'connections_correct' or (typ != 'diagram' and name in {'text_inside_bounds', 'no_overlap'})
        if state not in {'pass', 'fail', 'not_applicable'} or state == 'not_applicable' and not allowed_na:
            raise PolicyError('image-review-pending', '图片检查含未检查项目或错误豁免。', object_id=row['id'], check=name)
        if state == 'not_applicable' and not str((row.get('not_applicable_reasons') or {}).get(name, '')).strip():
            raise PolicyError('image-review-na-without-reason', '不适用项目须说明原因，不能借此跳过文字出框检查。', object_id=row['id'], check=name)
        if final and state == 'fail':
            raise PolicyError('image-visual-defect', '图片缺陷尚未修复，不得交付。', object_id=row['id'], check=name)
    if final and row.get('style_and_semantics_preserved') is not True:
        raise PolicyError('image-redraw-unverified', '修复图未与原图核对内容和原有风格。', object_id=row['id'])


def validate_initial(source, report):
    data = load(report)
    if data.get('source_sha256') != file_digest(source):
        raise PolicyError('initial-image-review-stale', '初检记录不是这份原始 Word。')
    actual = {o['id']: o for o in figures(source)}
    rows = indexed(data.get('objects'), actual)
    for ident, row in rows.items():
        if row.get('object_sha256') != actual[ident]['hash']:
            raise PolicyError('initial-image-review-stale', '初检图片哈希不匹配。', object_id=ident)
        validate_checks(row)
    return rows


def validate_final(source, candidate, initial_report, final_report, pages):
    initial = validate_initial(source, initial_report)
    data = load(final_report)
    if data.get('source_sha256') != file_digest(source) or data.get('candidate_sha256') != file_digest(candidate):
        raise PolicyError('final-image-review-stale', '图片终检记录不是当前原文与最终文件。')
    actual = {o['id']: o for o in figures(candidate)}
    if set(actual) != set(initial):
        raise PolicyError('image-inventory-changed', '图片对象增删或重组后须重新确认初检对应关系。')
    rows = indexed(data.get('objects'), actual)
    expected_pages = {p['name']: p['sha256'] for p in pages}
    for ident, row in rows.items():
        old = initial[ident]; obj = actual[ident]
        if row.get('object_sha256') != obj['hash'] or row.get('source_object_sha256') != old['object_sha256']:
            raise PolicyError('final-image-review-stale', '终检的原图或最终图哈希不匹配。', object_id=ident)
        if row.get('image_type') != old['image_type']:
            raise PolicyError('image-type-changed', '不得把有缺陷的架构图改称照片/装饰图而豁免检查。', object_id=ident)
        validate_checks(row, final=True)
        actual_failed = {k for k, v in old['checks'].items() if v == 'fail'}
        closed = row.get('resolved_initial_defects', [])
        if not isinstance(closed, list) or len(closed) != len(set(closed)) or set(closed) != actual_failed:
            raise PolicyError('image-issue-not-closed', '初检问题必须逐项修复并在最终图复查。', object_id=ident, unresolved=sorted(actual_failed - set(closed)))
        if actual_failed and not str(row.get('repair_note', '')).strip():
            raise PolicyError('image-repair-not-recorded', '已发现的图片问题缺少实际修复说明。', object_id=ident)
        if actual_failed & INTRINSIC and old['object_sha256'] == obj['hash']:
            raise PolicyError('image-defect-unchanged', '原图已有文字出框/重叠/连线错误/内嵌题注，但图像未变；只改报告不能算修复。', object_id=ident)
        if not actual_failed and old['object_sha256'] != obj['hash'] and not row.get('user_redesign_authorization'):
            raise PolicyError('image-unrequested-redesign', '正常原图被替换，且没有用户重新设计授权。', object_id=ident)
        refs = row.get('pages')
        if not isinstance(refs, list) or not refs or len(refs) != len({r.get('name') for r in refs}):
            raise PolicyError('image-final-page-missing', '每张图必须核对最新 Word 页面，不能只看脱离页面的重绘图。', object_id=ident)
        for ref in refs:
            if ref.get('name') not in expected_pages or expected_pages[ref['name']] != ref.get('sha256'):
                raise PolicyError('image-final-page-stale', '图片对应的最终页截图不匹配。', object_id=ident)
    return {'status': 'passed', 'figure_count': len(actual), 'initial_defects_closed': sum(sum(v == 'fail' for v in r['checks'].values()) for r in initial.values()),
            'note': '已验证图片检查记录的覆盖、哈希及闭环；像素语义和视觉结论由实际看图的 Agent 负责，不是 OCR 自动识别。'}


def main():
    ap = argparse.ArgumentParser(description='生成逐图检查清单或验证图片问题闭环')
    sub = ap.add_subparsers(dest='command', required=True)
    p = sub.add_parser('prepare'); p.add_argument('source', type=Path); p.add_argument('--out-dir', type=Path, required=True); p.add_argument('--json-out', type=Path, required=True)
    p = sub.add_parser('initial'); p.add_argument('source', type=Path); p.add_argument('review', type=Path); p.add_argument('--json-out', type=Path)
    p = sub.add_parser('final'); p.add_argument('source', type=Path); p.add_argument('candidate', type=Path); p.add_argument('initial_review', type=Path); p.add_argument('final_review', type=Path); p.add_argument('manifest', type=Path); p.add_argument('--json-out', type=Path)
    a = ap.parse_args()
    try:
        if a.command == 'prepare': result = make_review_template(a.source, a.out_dir)
        elif a.command == 'initial': result = {'status': 'passed', 'figure_count': len(validate_initial(a.source, a.review))}
        else: result = validate_final(a.source, a.candidate, a.initial_review, a.final_review, load(a.manifest)['render']['pages'])
    except (PolicyError, OSError, ValueError, KeyError, TypeError) as exc:
        result = {'status': 'failed', 'issues': [exc.as_issue() if isinstance(exc, PolicyError) else {'severity': 'error', 'code': 'image-review-invalid', 'message': str(exc)}]}
    payload = json.dumps(result, ensure_ascii=False, indent=2)
    if a.json_out: a.json_out.parent.mkdir(parents=True, exist_ok=True); a.json_out.write_text(payload+'\n', encoding='utf-8')
    print(payload); return 2 if result.get('status') == 'failed' else 0


if __name__ == '__main__': raise SystemExit(main())
