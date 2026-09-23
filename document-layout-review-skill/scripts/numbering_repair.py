#!/usr/bin/env python3
"""Apply template-authoritative numbering, keeping legacy native-footnote repair."""
from __future__ import annotations
import argparse
import json
import os
import tempfile
from pathlib import Path
from numbering_policy import (Doc, PolicyError, file_digest, inventory, load_object_plan,
                              normalize_document, visible)

HERE = Path(__file__).resolve().parent
DEFAULT_TEMPLATE = HERE.parent / 'assets' / 'default-template.docx'
DEFAULT_STYLE = HERE.parent / 'assets' / 'default-template-style.json'


def repair(source, out, template=None, style_json=None, numbering_plan=None, object_plan=None):
    source, out = Path(source), Path(out)
    template = Path(template) if template else DEFAULT_TEMPLATE
    style_json = Path(style_json) if style_json else (DEFAULT_STYLE if template.resolve() == DEFAULT_TEMPLATE.resolve() else None)
    if out.resolve() in {source.resolve(), template.resolve()}: raise PolicyError('in-place-write', '不得覆盖输入 Word 或模板。')
    # Validate semantic plan before any indices or file hashes change.
    choices = load_object_plan(object_plan, Doc(source))
    from numbering_core import Package, apply_plan, repair_footnotes, audit as legacy_audit
    out.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(dir=out.parent) as work:
        pre = Path(work) / 'notes.docx'; candidate = Path(work) / 'numbered.docx'
        pkg = Package(source)
        explicit = apply_plan(pkg, numbering_plan)
        footnote_changes = []
        repair_footnotes(pkg, footnote_changes, explicit)
        pkg.save(pre)
        current = Doc(pre); now = {o['id']: o for o in inventory(current)}
        rebound = []
        for ident, row in choices.items():
            if ident not in now or now[ident]['hash'] != row['object_sha256']:
                raise PolicyError('object-plan-stale', '脚注处理后图表内容发生变化，须重新定位。', object_id=ident)
            fresh = dict(row)
            if fresh.get('caption_paragraph') is not None:
                matches = [p for p in current.paragraphs if visible(p) == fresh['caption_text']]
                if len(matches) != 1: raise PolicyError('caption-plan-location', '脚注处理后题注锚点不唯一。')
                fresh['caption_paragraph'] = current.pindex[matches[0]]
            rebound.append(fresh)
        plan = {'version': 1, 'source_sha256': file_digest(pre), 'objects': rebound} if rebound else None
        result = normalize_document(pre, candidate, template, style_json, plan)
        checked = legacy_audit(candidate)
        if checked.get('status') != 'passed':
            raise PolicyError('native-numbering-audit-failed', '原生编号或脚注仍有问题，未写入输出。', issues=checked.get('issues', []))
        from layout_invariant_guard import snapshot_docx, changed_invariants
        changed = changed_invariants(snapshot_docx(source), snapshot_docx(candidate), 'automatic-numbering')
        if changed: raise PolicyError('numbering-collateral-edit', '编号修复改变了冻结对象。', invariants=changed)
        os.replace(candidate, out)
    result.update({'input': str(source), 'output': str(out), 'output_sha256': file_digest(out),
                   'scope': 'automatic-numbering', 'protected_invariants': 'passed',
                   'footnote_changes': footnote_changes, 'native_audit': checked})
    result['change_count'] += len(footnote_changes)
    return result


def main():
    ap = argparse.ArgumentParser(description='按实际模板修复标题编号；补齐图表题注及章内自动序列；修复原生脚注')
    ap.add_argument('input', type=Path); ap.add_argument('--out', type=Path, required=True)
    ap.add_argument('--template', type=Path); ap.add_argument('--template-style-json', type=Path)
    ap.add_argument('--numbering-plan', type=Path); ap.add_argument('--object-plan', type=Path)
    ap.add_argument('--json-out', type=Path); a = ap.parse_args()
    try: result = repair(a.input, a.out, a.template, a.template_style_json, a.numbering_plan, a.object_plan)
    except Exception as exc:
        result = {'status': 'blocked', 'output': None, 'changes': [], 'change_count': 0,
                  'issues': [exc.as_issue() if isinstance(exc, PolicyError) else {'severity': 'error', 'code': 'numbering-needs-review', 'message': str(exc)}]}
    payload = json.dumps(result, ensure_ascii=False, indent=2)
    if a.json_out: a.json_out.parent.mkdir(parents=True, exist_ok=True); a.json_out.write_text(payload+'\n', encoding='utf-8')
    print(payload); return 0 if result['status'] == 'passed' else 2


if __name__ == '__main__': raise SystemExit(main())
