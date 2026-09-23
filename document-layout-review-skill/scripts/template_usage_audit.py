#!/usr/bin/env python3
"""Audit ordinary body text against the active template and per-task rules."""
from __future__ import annotations

import argparse
import json
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from text_rules import (
    body_paragraph_items,
    TextRulesError, load_rules, load_profile, expected_body_props, effective_run_props,
    font_conflicts, rules_digest, EMPHASIS, run_font_slots,
)
from body_font_exceptions import preserved_body_fonts

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


def p_text(p: ET.Element) -> str:
    return "".join(t.text or "" for t in p.findall(".//w:t", NS))


def style_id(p: ET.Element) -> str | None:
    x = p.find("w:pPr/w:pStyle", NS)
    return x.get(qn("val")) if x is not None else None


def boolprop(rpr: ET.Element | None, tag: str) -> bool:
    if rpr is None:
        return False
    n = rpr.find(f"w:{tag}", NS)
    if n is None:
        return False
    v = n.get(qn("val"))
    return v not in {"0", "false", "off"}


def run_fonts(rpr: ET.Element | None) -> dict[str, str]:
    if rpr is None:
        return {}
    n = rpr.find("w:rFonts", NS)
    if n is None:
        return {}
    return {k: n.get(qn(k)) for k in FONT_KEYS if n.get(qn(k)) is not None}


def style_run_props(profile: dict, sid: str | None) -> dict:
    """Resolve run properties through basedOn inheritance."""
    out = dict(profile.get("doc_defaults", {}).get("run") or {})
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
                              rules: dict | None = None) -> dict:
    return font_conflicts(actual, expected, load_rules() if rules is None else rules)


def value_conflict(actual: str | None, expected: object | None) -> bool:
    return actual is not None and expected is not None and str(actual) != str(expected)


def character_style_conflicts(target_profile: dict, rpr: ET.Element | None, expected: dict, rules: dict | None = None) -> dict:
    if rpr is None:
        return {}
    rs = rpr.find("w:rStyle", NS)
    rsid = rs.get(qn("val")) if rs is not None else None
    if not rsid:
        return {}
    style_props = style_run_overrides(target_profile, rsid)
    bad: dict[str, object] = {}
    fonts = style_props.get("fonts") or {}
    font_bad = comparable_font_conflicts(fonts, expected.get("fonts") or {}, rules)
    if font_bad:
        bad["fonts"] = font_bad
    if value_conflict(style_props.get("size_half_points"), expected.get("size_half_points")):
        bad["size_half_points"] = {
            "actual": style_props.get("size_half_points"),
            "expected": expected.get("size_half_points"),
        }
    if value_conflict(
        style_props.get("size_cs_half_points"),
        expected.get("size_cs_half_points") or expected.get("size_half_points"),
    ):
        bad["size_cs_half_points"] = {
            "actual": style_props.get("size_cs_half_points"),
            "expected": expected.get("size_cs_half_points") or expected.get("size_half_points"),
        }
    return {"style_id": rsid, "conflicts": bad} if bad else {}


def direct_emphasis_issues(
    rpr: ET.Element | None,
    *,
    expected_bold: bool,
    expected_italic: bool,
    rules: dict | None = None,
) -> list[tuple[str, str]]:
    """Return direct bold/italic conflicts for ordinary body text.

    Character-style emphasis is intentionally not rejected here. This function
    only checks direct w:b/w:bCs/w:i/w:iCs properties.
    """
    if rpr is None:
        return []
    rules = load_rules() if rules is None else rules
    out: list[tuple[str, str]] = []
    if (rules["body"]["bold"] == "forbid" or (rules["body"]["bold"] == "template" and not expected_bold)) and (boolprop(rpr, "b") or boolprop(rpr, "bCs")):
        out.append(("body-run-bold-direct-format", "普通正文 run 存在直接加粗，覆盖了正文基准字重。"))
    if (rules["body"]["italic"] == "forbid" or (rules["body"]["italic"] == "template" and not expected_italic)) and (boolprop(rpr, "i") or boolprop(rpr, "iCs")):
        out.append(("body-run-italic-direct-format", "普通正文 run 存在直接斜体，覆盖了正文基准字形。"))
    return out


def audit(template: Path, target: Path, text_rules: Path | dict | None = None) -> dict:
    rules = load_rules(text_rules)
    template_profile = load_profile(template)
    target_profile = load_profile(target)
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

    with zipfile.ZipFile(target) as z:
        root = ET.fromstring(z.read("word/document.xml"))

    issues = []
    for idx, p in body_paragraph_items(root, template_profile, target_profile):
        sid = style_id(p)

        text = p_text(p).strip()
        if not text:
            continue

        # Paragraph-level run properties can make every run look bold/italic even
        # when individual runs have no rPr. Treat those as direct-format pollution.
        pmark_rpr = p.find("w:pPr/w:rPr", NS)
        for code, message in direct_emphasis_issues(
            pmark_rpr,
            expected_bold=expected_bold,
            expected_italic=expected_italic,
            rules=rules,
        ):
            issues.append(
                {
                    "severity": "error",
                    "code": code.replace("body-run-", "body-paragraph-"),
                    "paragraph": idx,
                    "run": 0,
                    "text": text[:120],
                    "message": message.replace("run", "段落默认文字"),
                }
            )

        for run_idx, r in enumerate(p.findall(".//w:r", NS), 1):
            rt = "".join(t.text or "" for t in r.findall(".//w:t", NS))
            if not rt:
                continue
            rpr = r.find("w:rPr", NS)
            actual = effective_run_props(target_profile, sid, rpr)
            actual_fonts = actual.get("fonts") or {}
            font_bad = font_conflicts(actual_fonts, expected_fonts, rules, run_font_slots(r))
            if font_bad:
                issues.append({
                    "severity": "error", "code": "body-run-font-override",
                    "paragraph": idx, "run": run_idx, "text": rt[:120],
                    "message": "正文实际字体不符合本次文字规则（含样式继承）。",
                    "actual": actual_fonts, "expected": expected_fonts, "conflicts": font_bad,
                })
            for key, exp, code in (
                ("size_half_points", expected_size, "body-run-size-override"),
                ("size_cs_half_points", expected_size_cs, "body-run-size-cs-override"),
            ):
                val = actual.get(key)
                if key == "size_cs_half_points" and val is None:
                    val = actual.get("size_half_points")
                if exp is not None and str(val) != str(exp):
                    issues.append({
                        "severity": "error", "code": code, "paragraph": idx,
                        "run": run_idx, "text": rt[:120],
                        "message": "正文实际字号不符合本次文字规则。",
                        "actual": val, "expected": str(exp),
                    })
            for key, tags in EMPHASIS.items():
                mode = rules["body"][key]
                if mode not in {"forbid", "require"}:
                    continue
                wanted = mode == "require"
                for tag, prop in zip(tags, (key, key + "_cs")):
                    if bool(actual.get(prop, False)) != wanted:
                        issues.append({
                            "severity": "error", "code": f"body-run-{tag}-rule",
                            "paragraph": idx, "run": run_idx, "text": rt[:120],
                            "message": f"正文 {tag} 不符合本次 {mode} 规则。",
                            "actual": bool(actual.get(prop, False)), "expected": wanted,
                        })

            # No coverage threshold: one stray direct-bold/direct-italic run is an
            # error whenever the body baseline itself is not bold/italic.
            for code, message in direct_emphasis_issues(
                rpr,
                expected_bold=expected_bold,
                expected_italic=expected_italic,
                rules=rules,
            ):
                issues.append(
                    {
                        "severity": "error",
                        "code": code,
                        "paragraph": idx,
                        "run": run_idx,
                        "text": rt[:120],
                        "message": message,
                    }
                )

    return {
        "template": str(template),
        "target": str(target),
        "status": "passed" if not issues else "failed",
        "body_style_id": body_sid,
        "expected_body_run": expected,
        "text_rules": rules,
        "text_rules_sha256": rules_digest(rules),
        "issues": issues,
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("template", type=Path)
    ap.add_argument("target", type=Path)
    ap.add_argument("--json-out", type=Path)
    ap.add_argument("--text-rules", type=Path, help="本次文字规则 JSON；不传则沿用默认规则")
    a = ap.parse_args()
    try:
        result = audit(a.template, a.target, a.text_rules)
    except TextRulesError as exc:
        result = {"status": "failed", "error": str(exc), "code": "text-rules-invalid"}
    payload = json.dumps(result, ensure_ascii=False, indent=2)
    if a.json_out:
        a.json_out.parent.mkdir(parents=True, exist_ok=True)
        a.json_out.write_text(payload + "\n", encoding="utf-8")
    print(payload)
    return 0 if result["status"] == "passed" else 2


if __name__ == "__main__":
    raise SystemExit(main())
