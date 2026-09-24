#!/usr/bin/env python3
"""Live negative-image test. Requires an authorized REAL vision worker.

Expected defects are used only by the evaluator and never sent to the worker.
This command is deliberately NOT part of the mock/unit test pass count.
"""
import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / 'scripts'))
import visual_evidence as ve


def intersects(a, b):
    return max(a[0], b[0]) < min(a[2], b[2]) and max(a[1], b[1]) < min(a[3], b[3])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--worker-config', type=Path, required=True)
    ap.add_argument('--out-dir', type=Path, required=True)
    args = ap.parse_args()
    image = ROOT / 'tests/fixtures/visual/right-column-overflow.png'
    task = ve.prepare([image], args.out_dir, {'fixture_sha256': ve.file_sha(image)}, phase='initial')
    receipt = ve.run(task, args.worker_config, args.out_dir)
    result = ve.validate_inspection(receipt)
    bundle = ve.validate_bundle(task)
    views = {v['id']: v for v in bundle['views']}
    # Ground truth from the supplied screenshot. NOT a production detector rule.
    expected = [[880, 463, 979, 483], [880, 497, 979, 519]]
    found = []
    for f in result['findings']:
        if f['check'] not in {'text_inside_bounds', 'no_overlap'}:
            continue
        x0, y0, x1, y1 = views[f['view_id']]['box']
        a, b, c, d = f['bbox']
        found.append([x0 + a * (x1-x0), y0 + b * (y1-y0), x0 + c * (x1-x0), y0 + d * (y1-y0)])
    ok = result['verdict'] == 'fail' and all(any(intersects(a, b) for b in found) for a in expected)
    report = {'status': 'passed' if ok else 'failed', 'kind': 'live-semantic-negative-test',
              'receipt': receipt, 'detected_boxes': found, 'expected_boxes': expected,
              'note': '只有真实视觉服务运行得到此结果，才能声称该服务识别了本负例。一次通过不是零漏检保证。'}
    ve.write(args.out_dir / 'live-regression-result.json', report)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if ok else 2


if __name__ == '__main__':
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(json.dumps({'status': 'blocked', 'message': str(exc)}, ensure_ascii=False))
        raise SystemExit(2)
