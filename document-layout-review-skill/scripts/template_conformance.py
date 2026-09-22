#!/usr/bin/env python3
"""Check whether a generated DOCX conforms to an active template and hard layout rules."""
from __future__ import annotations

import argparse
import json
from pathlib import Path

from template_style_profile import extract_profile


def _sig_map(profile):
    return {sid: item.get("signature") for sid, item in profile.get("styles", {}).items()}


def _is_technical_bid(profile: dict, template: Path) -> bool:
    if template.name == "default-technical-bid.docx":
        return True
    return any(bool(s.get("contains_technical_deviation")) for s in profile.get("sections", []))


def check(template: Path, target: Path):
    expected = extract_profile(template)
    actual = extract_profile(target)
    issues = []

    exp_hash = expected.get("package_hashes", {})
    act_hash = actual.get("package_hashes", {})
    for part in ("word/theme/theme1.xml",):
        if part in exp_hash and act_hash.get(part) != exp_hash.get(part):
            issues.append({"severity": "error", "code": "template-theme-changed", "message": f"目标文档未保留模板主题：{part}"})

    exp_styles = _sig_map(expected)
    act_styles = _sig_map(actual)
    required_style_ids = set()
    for role in expected.get("semantic_roles", {}).values():
        if isinstance(role, dict):
            sid = role.get("style_id")
            if sid:
                required_style_ids.add(sid)
            required_style_ids.update(role.get("style_ids", []) or [])
    required_style_ids.update(sid for sid, count in expected.get("usage", {}).get("table_styles", {}).items() if sid != "(direct)" and count)

    for sid in sorted(required_style_ids):
        if sid not in act_styles:
            issues.append({"severity": "error", "code": "template-style-missing", "style_id": sid, "message": f"目标文档缺少模板样式 {sid}"})
        elif act_styles[sid] != exp_styles.get(sid):
            issues.append({"severity": "error", "code": "template-style-mutated", "style_id": sid, "message": f"目标文档修改了模板样式 {sid} 的定义"})

    allowed_geometries = expected.get("section_geometries", [])
    if allowed_geometries:
        for idx, geo in enumerate(actual.get("section_geometries", []), 1):
            if geo not in allowed_geometries:
                issues.append({"severity": "error", "code": "template-section-geometry-mismatch", "section": idx, "message": "目标文档存在模板未定义的页面尺寸／边距／方向组合", "actual": geo})

    exp_roles = expected.get("semantic_roles", {})
    act_roles = actual.get("semantic_roles", {})
    for name, role in exp_roles.items():
        if not name.startswith("heading"):
            continue
        expected_sid = role.get("style_id") if isinstance(role, dict) else None
        actual_sid = act_roles.get(name, {}).get("style_id") if isinstance(act_roles.get(name), dict) else None
        if expected_sid and actual_sid and expected_sid != actual_sid:
            issues.append({"severity": "error", "code": "template-heading-style-mismatch", "role": name, "expected": expected_sid, "actual": actual_sid, "message": f"{name} 未使用模板对应样式"})

    # Table vertical alignment is a cell property, so style-ID conformance alone is insufficient.
    exp_table = expected.get("table_contract", {})
    act_table = actual.get("table_contract", {})
    if exp_table.get("all_cells_vertically_centered"):
        if not act_table.get("all_cells_vertically_centered"):
            issues.append({
                "severity": "error",
                "code": "template-table-vertical-alignment-mismatch",
                "message": "目标文档存在未上下居中的表格单元格；模板要求所有单元格上下居中。",
                "expected": "center",
                "actual": act_table.get("vertical_alignment_counts", {}),
            })

    # Explicit technical-bid policy: landscape belongs only to the technical-deviation section.
    if _is_technical_bid(expected, template):
        landscape = []
        for section in actual.get("sections", []):
            if section.get("orientation") != "landscape":
                continue
            landscape.append(section.get("index"))
            if not section.get("contains_technical_deviation"):
                issues.append({
                    "severity": "error",
                    "code": "technical-bid-landscape-outside-deviation-table",
                    "section": section.get("index"),
                    "message": "技术标只有“技术偏离表”所在节允许横向；该横向节不属于技术偏离表。",
                    "sample": section.get("text_sample", ""),
                })
        if len(landscape) > 1:
            issues.append({
                "severity": "error",
                "code": "technical-bid-multiple-landscape-sections",
                "message": "技术标最多只能有一个横向节，且该节必须用于技术偏离表。",
                "sections": landscape,
            })

    result = {
        "template": str(template),
        "target": str(target),
        "status": "passed" if not any(i["severity"] == "error" for i in issues) else "failed",
        "issues": issues,
        "expected_semantic_roles": exp_roles,
        "actual_semantic_roles": act_roles,
        "expected_table_contract": exp_table,
        "actual_table_contract": act_table,
        "actual_sections": actual.get("sections", []),
    }
    return result


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("template")
    ap.add_argument("target")
    ap.add_argument("--json-out")
    args = ap.parse_args()
    result = check(Path(args.template), Path(args.target))
    text = json.dumps(result, ensure_ascii=False, indent=2)
    if args.json_out:
        Path(args.json_out).write_text(text + "\n", encoding="utf-8")
    print(text)
    return 0 if result["status"] == "passed" else 2


if __name__ == "__main__":
    raise SystemExit(main())
