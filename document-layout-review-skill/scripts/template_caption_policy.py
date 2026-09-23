#!/usr/bin/env python3
"""Project caption baseline: figures below, tables above, native stable sequences.

The existing core already normalizes placement and chapter-local SEQ fields.
This adapter preserves every core issue and adds a strict field-dependency check.
Template styles control appearance; template samples never reverse placement.
"""
from __future__ import annotations
import argparse
import json
import os
import tempfile
from pathlib import Path
from template_format_contract import Q, NS, FormatContractError
from runtime_preflight import validate_profile
from caption_field_guard import PLACEMENT, audit as audit_caption_fields


def visible(p):
    return ''.join(t.text or '' for t in p.iter(Q('t'))).strip() if p is not None else ''


def adjacent(el, forward):
    n=el.getnext() if forward else el.getprevious()
    for _ in range(3):
        if n is None:return None
        if n.tag!=Q('p') or visible(n) or n.xpath('.//w:drawing|.//w:pict',namespaces=NS):return n
        n=n.getnext() if forward else n.getprevious()
    return None


def placement_rules(template, style_json=None, required=('figure','table')):
    """Compatibility entry: placement is a project invariant, not a template option.

    Deprecated caption_placement keys are deliberately ignored. Existing profiles
    stay readable, but cannot put a figure caption above or a table caption below.
    """
    from runtime_preflight import validate_docx
    validate_docx(template)
    if style_json is not None:
        validate_profile(template, style_json)
    requested = set(required)
    if requested - set(PLACEMENT):
        raise FormatContractError('caption-kind-invalid', '未知图表对象类型。')
    return {kind: PLACEMENT[kind] for kind in sorted(requested)}


def check_placement(doc, template, style_json=None, object_plan=None, repair=False):
    from numbering_policy import inventory, load_object_plan
    objects=inventory(doc)
    decisions=load_object_plan(object_plan,doc,for_repair=False) if object_plan else {}
    objects=[o for o in objects if not decisions.get(o['id'],{}).get('exempt')]
    policy=placement_rules(template,style_json,{o['kind'] for o in objects})
    issues=[]
    for obj in objects:
        cap=obj['caption'];owner=obj['element']
        if cap is None:continue  # missing-caption remains a blocking core audit error
        direction=policy[obj['kind']]
        if adjacent(owner,direction=='after') is cap:continue
        if cap.getparent() is not owner.getparent():
            raise FormatContractError('caption-placement-container','题注跨容器，不能移动并改变结构。',object_id=obj['id'])
        issue={'severity':'error','code':'caption-position','object_id':obj['id'],'expected':direction,
               'message':'题注位置必须为图下表上，不得由模板覆盖。'}
        issues.append(issue)
        if repair:
            (owner.addnext if direction=='after' else owner.addprevious)(cap)
            doc.changed.add('word/document.xml')
    if repair:doc.refresh()
    return issues


def audit_document(path, template, style_json=None, object_plan=None):
    from numbering_policy import audit_document as core_audit
    placement_rules(template, style_json)
    result = core_audit(path, template, style_json, object_plan)
    # Never suppress caption-position or any other core audit failure.
    field_check = audit_caption_fields(path)
    issues = list(result.get('issues', [])) + field_check['issues']
    return {**result, 'issues': issues,
            'status': 'passed' if result.get('status') == 'passed' and not issues else 'failed',
            'caption_field_contract': field_check, 'position_policy': dict(PLACEMENT)}


def normalize_document(source, out, template, style_json=None, object_plan=None):
    from numbering_policy import normalize_document as core_normalize, file_digest
    source, out = Path(source), Path(out)
    if out.resolve() in {source.resolve(), Path(template).resolve()}:
        raise FormatContractError('in-place-write', '编号输出不能覆盖输入或模板。')
    placement_rules(template, style_json)
    out.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(dir=out.parent) as td:
        numbered = Path(td) / 'numbered.docx'
        # The core removes stale STYLEREF prefixes and installs independent SEQ
        # fields. Do not move its correctly placed captions back to template errors.
        result = core_normalize(source, numbered, template, style_json, object_plan)
        bound = {'objects': result.get('object_decisions', [])}
        checked = audit_document(numbered, template, style_json, bound)
        if checked['status'] != 'passed':
            raise FormatContractError('caption-baseline-recheck-failed',
                                      '图下表上或自动编号稳定性复查失败。', issues=checked['issues'])
        os.replace(numbered, out)
    result.update(output=str(out), output_sha256=file_digest(out), audit=checked,
                  position_policy=dict(PLACEMENT))
    return result


def main():
    ap=argparse.ArgumentParser();ap.add_argument('action',choices=['audit','repair']);ap.add_argument('input',type=Path)
    ap.add_argument('--template',type=Path,required=True);ap.add_argument('--template-style-json',type=Path)
    ap.add_argument('--object-plan',type=Path);ap.add_argument('--out',type=Path);ap.add_argument('--json-out',type=Path);a=ap.parse_args()
    try:
        if a.action=='repair' and a.out is None:raise ValueError('repair 必须指定 --out')
        result=normalize_document(a.input,a.out,a.template,a.template_style_json,a.object_plan) if a.action=='repair' else audit_document(a.input,a.template,a.template_style_json,a.object_plan)
    except Exception as exc:
        result={'status':'failed','issues':[exc.as_issue() if hasattr(exc,'as_issue') else {'severity':'error','code':'caption-policy-error','message':str(exc)}]}
    data=json.dumps(result,ensure_ascii=False,indent=2)
    if a.json_out:
        if a.json_out.resolve() in {p.resolve() for p in (a.input,a.template,a.out,a.template_style_json,a.object_plan) if p is not None}:raise SystemExit('报告不能覆盖输入')
        a.json_out.parent.mkdir(parents=True,exist_ok=True);a.json_out.write_text(data+'\n',encoding='utf-8')
    print(data);return 0 if result['status']=='passed' else 2


if __name__=='__main__':raise SystemExit(main())
