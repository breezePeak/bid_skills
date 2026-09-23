#!/usr/bin/env python3
"""Repair body formatting using the same per-task rules as the audit."""
from __future__ import annotations

import argparse
import json
import zipfile
from pathlib import Path
from lxml import etree as ET

from text_rules import (
    body_paragraph_items,
    TextRulesError, load_rules, load_profile, expected_body_props, effective_run_props,
    font_conflicts, rules_digest, ensure_rpr, restore_emphasis, repair_effective_props,
)
from body_font_exceptions import (
    preserved_body_fonts,
    body_run_font_snapshot,
    clear_fonts_preserving_body_exception,
    restore_body_font_snapshot,
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
FONT_KEYS = (
    "ascii",
    "hAnsi",
    "eastAsia",
    "cs",
    "asciiTheme",
    "hAnsiTheme",
    "eastAsiaTheme",
    "csTheme",
)


def qn(n: str) -> str:
    return f"{{{W}}}{n}"


def text(p: ET.Element) -> str:
    return "".join(t.text or "" for t in p.findall(".//w:t", NS))


def sid(p: ET.Element) -> str | None:
    n = p.find("w:pPr/w:pStyle", NS)
    return n.get(qn("val")) if n is not None else None


def enabled(rpr: ET.Element | None, tag: str) -> bool:
    if rpr is None:
        return False
    n = rpr.find(f"w:{tag}", NS)
    if n is None:
        return False
    return n.get(qn("val")) not in {"0", "false", "off"}


def run_fonts(rpr: ET.Element | None) -> dict[str, str]:
    if rpr is None:
        return {}
    n = rpr.find("w:rFonts", NS)
    if n is None:
        return {}
    return {k: n.get(qn(k)) for k in FONT_KEYS if n.get(qn(k)) is not None}


def style_run_props(profile: dict, style_id: str | None) -> dict:
    out = dict(profile.get("doc_defaults", {}).get("run") or {})
    if not style_id:
        return out
    styles = profile.get("styles", {})
    chain: list[dict] = []
    seen: set[str] = set()
    cur = style_id
    while cur and cur not in seen:
        seen.add(cur)
        item = styles.get(cur)
        if not isinstance(item, dict):
            break
        chain.append(item)
        cur = item.get("based_on")
    for item in reversed(chain):
        out.update(item.get("run") or {})
    return out


def style_run_overrides(profile: dict, sid: str | None) -> dict:
    """Resolve only the properties explicitly supplied by a character-style chain."""
    out: dict = {}
    if not sid:
        return out
    styles = profile.get("styles", {})
    chain: list[dict] = []
    seen: set[str] = set()
    cur = sid
    while cur and cur not in seen:
        seen.add(cur)
        item = styles.get(cur)
        if not isinstance(item, dict):
            break
        chain.append(item)
        cur = item.get("based_on")
    for item in reversed(chain):
        out.update(item.get("run") or {})
    return out


def comparable_font_conflicts(actual: dict[str, str], expected: dict[str, str],
                              rules: dict | None = None) -> bool:
    return bool(font_conflicts(actual, expected, load_rules() if rules is None else rules))


def value_conflict(actual: str | None, expected: object | None) -> bool:
    return actual is not None and expected is not None and str(actual) != str(expected)


def character_style_conflicts(target_profile: dict, rpr: ET.Element | None, expected: dict, rules: dict | None = None) -> bool:
    if rpr is None:
        return False
    rs = rpr.find("w:rStyle", NS)
    rsid = rs.get(qn("val")) if rs is not None else None
    if not rsid:
        return False
    props = style_run_overrides(target_profile, rsid)
    if comparable_font_conflicts(props.get("fonts") or {}, expected.get("fonts") or {}, rules):
        return True
    if value_conflict(props.get("size_half_points"), expected.get("size_half_points")):
        return True
    if value_conflict(
        props.get("size_cs_half_points"),
        expected.get("size_cs_half_points") or expected.get("size_half_points"),
    ):
        return True
    return False


def remove_direct_emphasis(
    rpr: ET.Element | None,
    *,
    expected_bold: bool,
    expected_italic: bool,
    rules: dict | None = None,
) -> list[str]:
    if rpr is None:
        return []
    rules = load_rules() if rules is None else rules
    removed: list[str] = []
    if rules["body"]["bold"] == "forbid" or (rules["body"]["bold"] == "template" and not expected_bold):
        for tag in ("b", "bCs"):
            node = rpr.find(f"w:{tag}", NS)
            if node is not None and enabled(rpr, tag):
                rpr.remove(node)
                removed.append(tag)
    if rules["body"]["italic"] == "forbid" or (rules["body"]["italic"] == "template" and not expected_italic):
        for tag in ("i", "iCs"):
            node = rpr.find(f"w:{tag}", NS)
            if node is not None and enabled(rpr, tag):
                rpr.remove(node)
                removed.append(tag)
    return removed


def repair(template: Path, src: Path, out: Path, text_rules: Path | dict | None = None) -> dict:
    if src.resolve() == out.resolve() or template.resolve() == out.resolve():
        raise TextRulesError("输出路径必须与输入文档、模板不同。")
    rules = load_rules(text_rules)
    template_profile = load_profile(template)
    target_profile = load_profile(src)
    roles = template_profile.get("semantic_roles", {})
    body_sid = (roles.get("body") or {}).get("style_id")
    heading_sids = {
        v.get("style_id")
        for k, v in roles.items()
        if k.startswith("heading") and isinstance(v, dict) and v.get("style_id")
    }
    caption_role = roles.get("caption") if isinstance(roles.get("caption"), dict) else {}
    caption_sids = set(caption_role.get("style_ids", []) or [])
    if caption_role.get("style_id"):
        caption_sids.add(caption_role["style_id"])

    expected = expected_body_props(template_profile, rules)
    expected_fonts = expected.get("fonts") or {}
    expected_size = expected.get("size_half_points")
    expected_size_cs = expected.get("size_cs_half_points") or expected_size
    expected_bold = expected.get("bold") is True
    expected_italic = expected.get("italic") is True

    with zipfile.ZipFile(src) as z:
        root = ET.fromstring(z.read("word/document.xml"))

    changes = []
    for idx, p in body_paragraph_items(root, template_profile, target_profile):
        psid = sid(p)

        raw = text(p).strip()
        if not raw:
            continue

        # Paragraph-level run properties can impose direct bold/italic on all runs.
        pmark_rpr = p.find("w:pPr/w:rPr", NS)
        pmark_removed = remove_direct_emphasis(
            pmark_rpr,
            expected_bold=expected_bold,
            expected_italic=expected_italic,
            rules=rules,
        )
        if pmark_removed:
            changes.append(
                {
                    "paragraph": idx,
                    "run": 0,
                    "removed": pmark_removed,
                    "text": raw[:120],
                    "scope": "paragraph-default-run-properties",
                }
            )
        if pmark_rpr is not None and len(pmark_rpr) == 0:
            ppr = p.find("w:pPr", NS)
            if ppr is not None:
                ppr.remove(pmark_rpr)

        for run_idx, r in enumerate(p.findall(".//w:r", NS), 1):
            rt = "".join(t.text or "" for t in r.findall(".//w:t", NS))
            if not rt:
                continue
            before = ET.tostring(r)
            snapshot = effective_run_props(target_profile, psid, r.find("w:rPr", NS))
            kept_fonts = body_run_font_snapshot(target_profile, psid, r, rules)
            rpr = ensure_rpr(r)
            removed: list[str] = []

            actual_fonts = run_fonts(rpr)
            if actual_fonts and comparable_font_conflicts(actual_fonts, expected_fonts, rules):
                if clear_fonts_preserving_body_exception(rpr, rules):
                    removed.append("rFonts")
            for tag, exp in (("sz", expected_size), ("szCs", expected_size_cs)):
                node = rpr.find(f"w:{tag}", NS)
                val = node.get(qn("val")) if node is not None else None
                if node is not None and value_conflict(val, exp):
                    rpr.remove(node)
                    removed.append(tag)

            # Keep unrelated character-style properties; repair their conflicting
            # font/size locally rather than removing the entire character style.
            removed.extend(remove_direct_emphasis(
                rpr, expected_bold=expected_bold, expected_italic=expected_italic, rules=rules,
            ))
            if removed:
                restore_body_font_snapshot(r, kept_fonts)
            if any(rules["body"][key] == "preserve" for key in ("bold", "italic")):
                restore_emphasis(rpr, snapshot, rules)
            updated = repair_effective_props(target_profile, psid, r, expected, rules)
            if len(rpr) == 0:
                r.remove(rpr)
            if ET.tostring(r) != before:
                changes.append({
                    "paragraph": idx, "run": run_idx, "removed": removed,
                    "updated": updated, "text": rt[:120], "scope": "run",
                })

    xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    out.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(src) as z, zipfile.ZipFile(out, "w") as dst:
        for info in z.infolist():
            dst.writestr(info, xml if info.filename == "word/document.xml" else z.read(info.filename))

    return {
        "input": str(src),
        "output": str(out),
        "change_count": len(changes),
        "expected_body_run": expected,
        "text_rules": rules,
        "text_rules_sha256": rules_digest(rules),
        "changes": changes,
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("template", type=Path)
    ap.add_argument("input", type=Path)
    ap.add_argument("--out", type=Path, required=True)
    ap.add_argument("--json-out", type=Path)
    ap.add_argument("--text-rules", type=Path, help="本次文字规则 JSON")
    a = ap.parse_args()
    try:
        result = repair(a.template, a.input, a.out, a.text_rules)
    except TextRulesError as exc:
        result = {"status": "failed", "error": str(exc), "code": "text-rules-invalid"}
    payload = json.dumps(result, ensure_ascii=False, indent=2)
    if a.json_out:
        a.json_out.parent.mkdir(parents=True, exist_ok=True)
        a.json_out.write_text(payload + "\n", encoding="utf-8")
    print(payload)
    return 2 if result.get("status") == "failed" else 0


if __name__ == "__main__":
    raise SystemExit(main())
