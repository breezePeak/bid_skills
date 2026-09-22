#!/usr/bin/env python3
"""Repair body-format drift at run granularity.

Any visible ordinary-body run whose explicit font/size conflicts with the template
baseline is repaired locally. Direct bold/italic are also removed whenever the body
baseline itself is not bold/italic. Intentional emphasis should be expressed through
an allowed character style instead of stray direct formatting.
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


def comparable_font_conflicts(actual: dict[str, str], expected: dict[str, str]) -> bool:
    for key in FONT_KEYS:
        if key in actual and key in expected and actual[key] != expected[key]:
            return True
    return False


def value_conflict(actual: str | None, expected: object | None) -> bool:
    return actual is not None and expected is not None and str(actual) != str(expected)


def character_style_conflicts(target_profile: dict, rpr: ET.Element | None, expected: dict) -> bool:
    if rpr is None:
        return False
    rs = rpr.find("w:rStyle", NS)
    rsid = rs.get(qn("val")) if rs is not None else None
    if not rsid:
        return False
    props = style_run_overrides(target_profile, rsid)
    if comparable_font_conflicts(props.get("fonts") or {}, expected.get("fonts") or {}):
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
) -> list[str]:
    if rpr is None:
        return []
    removed: list[str] = []
    if not expected_bold:
        for tag in ("b", "bCs"):
            node = rpr.find(f"w:{tag}", NS)
            if node is not None and enabled(rpr, tag):
                rpr.remove(node)
                removed.append(tag)
    if not expected_italic:
        for tag in ("i", "iCs"):
            node = rpr.find(f"w:{tag}", NS)
            if node is not None and enabled(rpr, tag):
                rpr.remove(node)
                removed.append(tag)
    return removed


def repair(template: Path, src: Path, out: Path) -> dict:
    template_profile = extract_profile(template)
    target_profile = extract_profile(src)
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
    expected_bold = expected.get("bold") is True
    expected_italic = expected.get("italic") is True

    with zipfile.ZipFile(src) as z:
        root = ET.fromstring(z.read("word/document.xml"))

    table_paras = {
        id(p)
        for tbl in root.findall(".//w:tbl", NS)
        for p in tbl.findall(".//w:p", NS)
    }

    changes = []
    in_body = False
    for idx, p in enumerate(root.findall(".//w:p", NS), 1):
        psid = sid(p)
        if psid in heading_sids:
            in_body = True
            continue
        if psid == body_sid:
            in_body = True
        if not in_body or id(p) in table_paras or psid in caption_sids:
            continue

        raw = text(p).strip()
        if not raw:
            continue

        # Paragraph-level run properties can impose direct bold/italic on all runs.
        pmark_rpr = p.find("w:pPr/w:rPr", NS)
        pmark_removed = remove_direct_emphasis(
            pmark_rpr,
            expected_bold=expected_bold,
            expected_italic=expected_italic,
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
            rpr = r.find("w:rPr", NS)
            if rpr is None:
                continue
            removed: list[str] = []

            actual_fonts = run_fonts(rpr)
            if actual_fonts and comparable_font_conflicts(actual_fonts, expected_fonts):
                node = rpr.find("w:rFonts", NS)
                if node is not None:
                    rpr.remove(node)
                    removed.append("rFonts")

            for tag, exp in (("sz", expected_size), ("szCs", expected_size_cs)):
                node = rpr.find(f"w:{tag}", NS)
                val = node.get(qn("val")) if node is not None else None
                if node is not None and value_conflict(val, exp):
                    rpr.remove(node)
                    removed.append(tag)

            # A character style that changes font/size away from the body baseline
            # is formatting pollution. Character styles that only provide semantic
            # emphasis (for example bold) are preserved.
            if character_style_conflicts(target_profile, rpr, expected):
                node = rpr.find("w:rStyle", NS)
                if node is not None:
                    rpr.remove(node)
                    removed.append("rStyle")

            # Direct bold/italic no longer use a paragraph coverage threshold.
            # One stray run is enough to produce the visible defect shown by users.
            removed.extend(
                remove_direct_emphasis(
                    rpr,
                    expected_bold=expected_bold,
                    expected_italic=expected_italic,
                )
            )

            if removed:
                changes.append(
                    {
                        "paragraph": idx,
                        "run": run_idx,
                        "removed": removed,
                        "text": rt[:120],
                        "scope": "run",
                    }
                )
            if len(rpr) == 0:
                r.remove(rpr)

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
        "changes": changes,
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("template", type=Path)
    ap.add_argument("input", type=Path)
    ap.add_argument("--out", type=Path, required=True)
    ap.add_argument("--json-out", type=Path)
    a = ap.parse_args()
    result = repair(a.template, a.input, a.out)
    payload = json.dumps(result, ensure_ascii=False, indent=2)
    if a.json_out:
        a.json_out.parent.mkdir(parents=True, exist_ok=True)
        a.json_out.write_text(payload + "\n", encoding="utf-8")
    print(payload)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
