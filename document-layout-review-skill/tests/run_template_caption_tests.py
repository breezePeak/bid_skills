#!/usr/bin/env python3
"""Run this patch's targeted cases; do not claim repository-wide coverage."""
import argparse
import json
import sys
import unittest
from pathlib import Path

MODULES = (
    'test_visual_preservation',
    'test_numbering_policy',
    'test_template_contract_extra',
    'test_figure_gate',
    'test_release_gate',
    'test_pipeline_boundary',
    'test_template_authority',
)

def main():
    ap=argparse.ArgumentParser();ap.add_argument('--json-out',type=Path);args=ap.parse_args()
    sys.path.insert(0,str(Path(__file__).resolve().parent))
    loader=unittest.TestLoader();suite=unittest.TestSuite(loader.loadTestsFromName(name) for name in MODULES)
    result=unittest.TextTestRunner(verbosity=2).run(suite)
    report={'scope':'targeted patch regression, not whole repository or real Word integration',
            'modules':list(MODULES),'tests_run':result.testsRun,'failures':len(result.failures),
            'errors':len(result.errors),'skipped':len(result.skipped),
            'status':'passed' if result.wasSuccessful() else 'failed'}
    if args.json_out:
        args.json_out.parent.mkdir(parents=True,exist_ok=True)
        args.json_out.write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps(report,ensure_ascii=False,indent=2));return 0 if result.wasSuccessful() else 1

if __name__=='__main__':raise SystemExit(main())
