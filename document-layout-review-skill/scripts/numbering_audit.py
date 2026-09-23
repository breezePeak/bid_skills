#!/usr/bin/env python3
"""Audit final native numbering against the same active template as repair."""
from __future__ import annotations
import argparse
import json
from pathlib import Path
from numbering_policy import audit_document, PolicyError


def audit(source, template=None, style_json=None, object_plan=None):
    from numbering_repair import DEFAULT_TEMPLATE, DEFAULT_STYLE
    from numbering_core import audit as native_audit
    template = Path(template) if template else DEFAULT_TEMPLATE
    style_json = Path(style_json) if style_json else (DEFAULT_STYLE if template.resolve() == DEFAULT_TEMPLATE.resolve() else None)
    policy = audit_document(Path(source), template, style_json, object_plan)
    native = native_audit(Path(source))
    issues = policy['issues'] + native.get('issues', [])
    return {'status': 'passed' if not issues else 'failed', 'issues': issues,
            'counts': native.get('counts', {}), 'policy': policy, 'native': native,
            'visual_review_required': True, 'field_update_required': True}


def main():
    ap = argparse.ArgumentParser(); ap.add_argument('input', type=Path)
    ap.add_argument('--template', type=Path); ap.add_argument('--template-style-json', type=Path)
    ap.add_argument('--object-plan', type=Path); ap.add_argument('--json-out', type=Path)
    a = ap.parse_args()
    try: result = audit(a.input, a.template, a.template_style_json, a.object_plan)
    except Exception as exc: result = {'status': 'failed', 'issues': [exc.as_issue() if isinstance(exc, PolicyError) else {'code': 'numbering-audit-error', 'message': str(exc), 'severity': 'error'}]}
    payload = json.dumps(result, ensure_ascii=False, indent=2)
    if a.json_out: a.json_out.parent.mkdir(parents=True, exist_ok=True); a.json_out.write_text(payload+'\n', encoding='utf-8')
    print(payload); return 0 if result['status'] == 'passed' else 2


if __name__ == '__main__': raise SystemExit(main())
