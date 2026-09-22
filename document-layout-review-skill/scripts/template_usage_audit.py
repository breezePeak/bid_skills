#!/usr/bin/env python3
"""Audit actual body formatting at run granularity.

Font and size are structural formatting, not semantic emphasis. A single visible
body run whose explicit font/size conflicts with the body baseline is a release-
blocking error. Bold/italic keep the older bulk threshold because short local
emphasis may be intentional.
"""
from __future__ import annotations

import argparse
import json
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from template_style_profile import extract_profile

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


def comparable_font_conflicts(actual: dict[str, str], expected: dict[str, str]) -> dict:
    """Return only conflicts we can determine without resolving theme fonts.

    If one side uses a concrete font and the other only a theme token, do not guess;
    template_style_enforce already removes direct rFonts from known body paragraphs.
    """
    bad = {}
    for key in FONT_KEYS:
        if key in actual and key in expected and actual[key] != expected[key]:
            bad[key] = {"actual": actual[key], "expected": expected[key]}
    return bad


def value_conflict(actual: str | None, expected: object | None) -> bool:
    return actual is not None and expected is not None and str(actual) != str(expected)


def character_style_conflicts(target_profile: dict, rpr: ET.Element | None, expected: dict) -> dict:
    if rpr is None:
        return {}
    rs = rpr.find("w:rStyle", NS)
    rsid = rs.get(qn("val")) if rs is not None else None
    if not rsid:
        return {}
    style_props = style_run_overrides(target_profile, rsid)
    bad: dict[str, object] = {}
    fonts = style_props.get("fonts") or {}
    font_bad = comparable_font_conflicts(fonts, expected.get("fonts") or {})
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


def audit(template: Path, target: Path) -> dict:
    template_profile = extract_profile(template)
    target_profile = extract_profile(target)
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

    expected = style_run_props(template_profile, body_sid)
    expected_fonts = expected.get("fonts") or {}
    expected_size = expected.get("size_half_points")
    expected_size_cs = expected.get("size_cs_half_points") or expected_size

    with zipfile.ZipFile(target) as z:
        root = ET.fromstring(z.read("word/document.xml"))

    table_paras = {
        id(p)
        for tbl in root.findall(".//w:tbl", NS)
        for p in tbl.findall(".//w:p", NS)
    }

    issues = []
    in_body = False
    for idx, p in enumerate(root.findall(".//w:p", NS), 1):
        sid = style_id(p)
        if sid in heading_sids:
            in_body = True
            continue
        if sid == body_sid:
            in_body = True
        if not in_body or id(p) in table_paras or sid in caption_sids:
            continue

        text = p_text(p).strip()
        if not text:
            continue

        total = 0
        bold = 0
        italic = 0
        for run_idx, r in enumerate(p.findall(".//w:r", NS), 1):
            rt = "".join(t.text or "" for t in r.findall(".//w:t", NS))
            n = len(rt)
            if not n:
                continue
            total += n
            rpr = r.find("w:rPr", NS)
            if boolprop(rpr, "b"):
                bold += n
            if boolprop(rpr, "i"):
                italic += n
            if rpr is None:
                continue

            actual_fonts = run_fonts(rpr)
            font_bad = comparable_font_conflicts(actual_fonts, expected_fonts)
            if font_bad:
                issues.append(
                    {
                        "severity": "error",
                        "code": "body-run-font-override",
                        "paragraph": idx,
                        "run": run_idx,
                        "text": rt[:120],
                        "message": "普通正文局部 run 使用了与正文基准冲突的显式字体。",
                        "actual": actual_fonts,
                        "expected": expected_fonts,
                        "conflicts": font_bad,
                    }
                )

            for tag, exp, code in (
                ("sz", expected_size, "body-run-size-override"),
                ("szCs", expected_size_cs, "body-run-size-cs-override"),
            ):
                node = rpr.find(f"w:{tag}", NS)
                val = node.get(qn("val")) if node is not None else None
                if value_conflict(val, exp):
                    issues.append(
                        {
                            "severity": "error",
                            "code": code,
                            "paragraph": idx,
                            "run": run_idx,
                            "text": rt[:120],
                            "message": "普通正文局部 run 使用了与正文基准冲突的显式字号。",
                            "actual": val,
                            "expected": str(exp),
                        }
                    )

            char_bad = character_style_conflicts(target_profile, rpr, expected)
            if char_bad:
                issues.append(
                    {
                        "severity": "error",
                        "code": "body-run-character-style-font-size-override",
                        "paragraph": idx,
                        "run": run_idx,
                        "text": rt[:120],
                        "message": "普通正文局部 run 的字符样式覆盖了正文基准字体或字号。",
                        **char_bad,
                    }
                )

        # 只有加粗/斜体继续使用覆盖率阈值；字体/字号已经逐 run 零容忍检查。
        if total >= 12:
            if bold / total >= 0.60:
                issues.append(
                    {
                        "severity": "error",
                        "code": "body-bulk-bold-direct-format",
                        "paragraph": idx,
                        "coverage": round(bold / total, 3),
                        "text": text[:160],
                        "message": "普通正文大面积使用直接加粗，覆盖了模板正文样式。",
                    }
                )
            if italic / total >= 0.60:
                issues.append(
                    {
                        "severity": "error",
                        "code": "body-bulk-italic-direct-format",
                        "paragraph": idx,
                        "coverage": round(italic / total, 3),
                        "text": text[:160],
                        "message": "普通正文大面积使用直接斜体，覆盖了模板正文样式。",
                    }
                )

    return {
        "template": str(template),
        "target": str(target),
        "status": "passed" if not issues else "failed",
        "body_style_id": body_sid,
        "expected_body_run": expected,
        "issues": issues,
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("template", type=Path)
    ap.add_argument("target", type=Path)
    ap.add_argument("--json-out", type=Path)
    a = ap.parse_args()
    result = audit(a.template, a.target)
    payload = json.dumps(result, ensure_ascii=False, indent=2)
    if a.json_out:
        a.json_out.parent.mkdir(parents=True, exist_ok=True)
        a.json_out.write_text(payload + "\n", encoding="utf-8")
    print(payload)
    return 0 if result["status"] == "passed" else 2


if __name__ == "__main__":
    raise SystemExit(main())
