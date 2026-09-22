#!/usr/bin/env python3
"""Extract a deterministic formatting profile from a DOCX template.

The profile is designed for document-generation agents. It captures the template's
actual Word style definitions and page geometry instead of asking the model to infer
formatting from screenshots alone.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"w": W, "r": R}
A = lambda name: f"{{{W}}}{name}"


def _val(node: ET.Element | None, name: str = "val") -> str | None:
    return None if node is None else node.get(A(name))


def _bool_prop(parent: ET.Element | None, tag: str) -> bool | None:
    if parent is None:
        return None
    node = parent.find(f"w:{tag}", NS)
    if node is None:
        return None
    value = _val(node)
    return value not in {"0", "false", "off"}


def _props(node: ET.Element | None, kind: str) -> dict[str, Any]:
    if node is None:
        return {}
    out: dict[str, Any] = {}
    if kind == "run":
        fonts = node.find("w:rFonts", NS)
        if fonts is not None:
            out["fonts"] = {k: fonts.get(A(k)) for k in ("ascii", "hAnsi", "eastAsia", "cs", "asciiTheme", "hAnsiTheme", "eastAsiaTheme", "csTheme") if fonts.get(A(k)) is not None}
        for tag, key in (("sz", "size_half_points"), ("szCs", "size_cs_half_points"), ("color", "color"), ("highlight", "highlight"), ("vertAlign", "vertical_align")):
            n = node.find(f"w:{tag}", NS)
            if n is not None and _val(n) is not None:
                out[key] = _val(n)
        for tag, key in (("b", "bold"), ("i", "italic"), ("u", "underline"), ("strike", "strike"), ("caps", "caps"), ("smallCaps", "small_caps")):
            value = _bool_prop(node, tag)
            if value is not None:
                out[key] = value
    elif kind == "paragraph":
        jc = node.find("w:jc", NS)
        if jc is not None and _val(jc) is not None:
            out["alignment"] = _val(jc)
        spacing = node.find("w:spacing", NS)
        if spacing is not None:
            vals = {k: spacing.get(A(k)) for k in ("before", "after", "line", "lineRule", "beforeLines", "afterLines") if spacing.get(A(k)) is not None}
            if vals:
                out["spacing"] = vals
        ind = node.find("w:ind", NS)
        if ind is not None:
            vals = {k: ind.get(A(k)) for k in ("left", "right", "firstLine", "hanging", "firstLineChars", "hangingChars") if ind.get(A(k)) is not None}
            if vals:
                out["indent"] = vals
        for tag, key in (("keepNext", "keep_next"), ("keepLines", "keep_lines"), ("pageBreakBefore", "page_break_before"), ("widowControl", "widow_control")):
            value = _bool_prop(node, tag)
            if value is not None:
                out[key] = value
        ol = node.find("w:outlineLvl", NS)
        if ol is not None and _val(ol) is not None:
            out["outline_level"] = int(_val(ol))
    elif kind == "table":
        style = node.find("w:tblStyle", NS)
        if style is not None and _val(style) is not None:
            out["style_id"] = _val(style)
        width = node.find("w:tblW", NS)
        if width is not None:
            out["width"] = {"value": width.get(A("w")), "type": width.get(A("type"))}
        layout = node.find("w:tblLayout", NS)
        if layout is not None and _val(layout, "type") is not None:
            out["layout"] = _val(layout, "type")
        look = node.find("w:tblLook", NS)
        if look is not None:
            out["look"] = {k: look.get(A(k)) for k in ("val", "firstRow", "lastRow", "firstColumn", "lastColumn", "noHBand", "noVBand") if look.get(A(k)) is not None}
        cell_mar = node.find("w:tblCellMar", NS)
        if cell_mar is not None:
            mar: dict[str, Any] = {}
            for side in ("top", "left", "bottom", "right", "start", "end"):
                s = cell_mar.find(f"w:{side}", NS)
                if s is not None:
                    mar[side] = {"value": s.get(A("w")), "type": s.get(A("type"))}
            if mar:
                out["cell_margins"] = mar
    return out


def _style_signature(style: dict[str, Any]) -> str:
    encoded = json.dumps(style, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return "sha256:" + hashlib.sha256(encoded).hexdigest()


def _paragraph_direct_signature(p: ET.Element) -> dict[str, Any]:
    ppr = p.find("w:pPr", NS)
    p_props = _props(ppr, "paragraph")
    runs = []
    for r in p.findall("w:r", NS):
        text = "".join(t.text or "" for t in r.findall(".//w:t", NS)).strip()
        rpr = r.find("w:rPr", NS)
        if text:
            runs.append(_props(rpr, "run"))
    run_sig: dict[str, Any] = {}
    if runs:
        key_counts = Counter(json.dumps(v, ensure_ascii=False, sort_keys=True) for v in runs)
        run_sig = json.loads(key_counts.most_common(1)[0][0])
    return {"paragraph": p_props, "run": run_sig}


def _text(p: ET.Element) -> str:
    return "".join(t.text or "" for t in p.findall(".//w:t", NS)).strip()


def _section_geometry(sect: ET.Element) -> dict[str, Any]:
    out: dict[str, Any] = {}
    size = sect.find("w:pgSz", NS)
    if size is not None:
        out["page_size"] = {k: size.get(A(k)) for k in ("w", "h", "orient", "code") if size.get(A(k)) is not None}
    mar = sect.find("w:pgMar", NS)
    if mar is not None:
        out["margins"] = {k: mar.get(A(k)) for k in ("top", "right", "bottom", "left", "header", "footer", "gutter") if mar.get(A(k)) is not None}
    cols = sect.find("w:cols", NS)
    if cols is not None:
        out["columns"] = {k: cols.get(A(k)) for k in ("num", "space", "equalWidth") if cols.get(A(k)) is not None}
    grid = sect.find("w:docGrid", NS)
    if grid is not None:
        out["doc_grid"] = {k: grid.get(A(k)) for k in ("type", "linePitch", "charSpace") if grid.get(A(k)) is not None}
    return out


def _node_text(node: ET.Element) -> str:
    return "".join(t.text or "" for t in node.findall(".//w:t", NS)).strip()


def _orientation(sect: ET.Element) -> str:
    size = sect.find("w:pgSz", NS)
    if size is None:
        return "portrait"
    orient = size.get(A("orient"))
    try:
        width = int(size.get(A("w")) or "0")
        height = int(size.get(A("h")) or "0")
    except ValueError:
        width = height = 0
    return "landscape" if orient == "landscape" or (width > 0 and height > 0 and width > height) else "portrait"


def _section_records(body: ET.Element) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    current: list[ET.Element] = []

    def emit(sect: ET.Element) -> None:
        nonlocal current
        text = " ".join(_node_text(el) for el in current if _node_text(el))
        records.append({
            "index": len(records) + 1,
            "orientation": _orientation(sect),
            "geometry": _section_geometry(sect),
            "contains_technical_deviation": "技术偏离表" in text.replace(" ", ""),
            "text_sample": text[:240],
        })
        current = []

    for child in list(body):
        if child.tag == A("sectPr"):
            emit(child)
            continue
        current.append(child)
        if child.tag == A("p"):
            sect = child.find("w:pPr/w:sectPr", NS)
            if sect is not None:
                emit(sect)
    return records


def _table_contract(body: ET.Element | None) -> dict[str, Any]:
    counts: Counter[str] = Counter()
    total = 0
    if body is not None:
        for tc in body.findall(".//w:tc", NS):
            total += 1
            valign = tc.find("w:tcPr/w:vAlign", NS)
            value = _val(valign) if valign is not None else None
            counts[value or "(unset)"] += 1
    dominant = counts.most_common(1)[0][0] if counts else None
    return {
        "cell_count": total,
        "vertical_alignment_counts": dict(counts),
        "dominant_vertical_alignment": dominant,
        "all_cells_vertically_centered": total > 0 and counts.get("center", 0) == total,
    }


def _canonical_xml_digest(data: bytes) -> str:
    root = ET.fromstring(data)
    def walk(e: ET.Element) -> Any:
        attrs = sorted((k, v) for k, v in e.attrib.items())
        text = (e.text or "").strip()
        return [e.tag, attrs, text, [walk(c) for c in list(e)]]
    encoded = json.dumps(walk(root), ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return "sha256:" + hashlib.sha256(encoded).hexdigest()


def extract_profile(path: Path) -> dict[str, Any]:
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        styles_root = ET.fromstring(z.read("word/styles.xml"))
        document_root = ET.fromstring(z.read("word/document.xml"))
        styles: dict[str, Any] = {}
        outline_to_style: dict[str, str] = {}
        name_to_style: dict[str, str] = {}
        for s in styles_root.findall("w:style", NS):
            sid = s.get(A("styleId")) or ""
            item: dict[str, Any] = {
                "id": sid,
                "type": s.get(A("type")),
                "name": _val(s.find("w:name", NS)) or "",
                "based_on": _val(s.find("w:basedOn", NS)),
                "next": _val(s.find("w:next", NS)),
                "link": _val(s.find("w:link", NS)),
                "paragraph": _props(s.find("w:pPr", NS), "paragraph"),
                "run": _props(s.find("w:rPr", NS), "run"),
                "table": _props(s.find("w:tblPr", NS), "table"),
                "qformat": s.find("w:qFormat", NS) is not None,
                "default": s.get(A("default")) in {"1", "true", "on"},
            }
            item["signature"] = _style_signature({k: v for k, v in item.items() if k != "signature"})
            styles[sid] = item
            if item["name"]:
                name_to_style[str(item["name"])] = sid
            ol = item["paragraph"].get("outline_level")
            if ol is not None and item["type"] == "paragraph":
                outline_to_style[str(int(ol) + 1)] = sid

        doc_defaults: dict[str, Any] = {}
        dd = styles_root.find("w:docDefaults", NS)
        if dd is not None:
            doc_defaults["run"] = _props(dd.find("w:rPrDefault/w:rPr", NS), "run")
            doc_defaults["paragraph"] = _props(dd.find("w:pPrDefault/w:pPr", NS), "paragraph")

        body = document_root.find("w:body", NS)
        paragraph_usage: Counter[str] = Counter()
        table_paragraph_usage: Counter[str] = Counter()
        direct_outside: Counter[str] = Counter()
        direct_inside: Counter[str] = Counter()
        if body is not None:
            table_paras = set(id(p) for tbl in body.findall(".//w:tbl", NS) for p in tbl.findall(".//w:p", NS))
            for p in body.findall(".//w:p", NS):
                if not _text(p):
                    continue
                ps = p.find("w:pPr/w:pStyle", NS)
                sid = _val(ps) or "(direct)"
                key = json.dumps(_paragraph_direct_signature(p), ensure_ascii=False, sort_keys=True)
                if id(p) in table_paras:
                    table_paragraph_usage[sid] += 1
                    direct_inside[key] += 1
                else:
                    paragraph_usage[sid] += 1
                    direct_outside[key] += 1

        table_usage: Counter[str] = Counter()
        if body is not None:
            for tbl in body.findall(".//w:tbl", NS):
                sid = _val(tbl.find("w:tblPr/w:tblStyle", NS)) or "(direct)"
                table_usage[sid] += 1

        section_geometries = []
        sections: list[dict[str, Any]] = []
        table_contract = _table_contract(body)
        if body is not None:
            sections = _section_records(body)
            for section in sections:
                geo = section.get("geometry") or {}
                if geo and geo not in section_geometries:
                    section_geometries.append(geo)

        semantic_roles: dict[str, Any] = {}
        for level, sid in sorted(outline_to_style.items(), key=lambda kv: int(kv[0])):
            semantic_roles[f"heading{level}"] = {"style_id": sid, "source": "outline-level"}

        caption_candidates = []
        for sid, item in styles.items():
            if item.get("type") != "paragraph":
                continue
            name = str(item.get("name", ""))
            if re.search(r"caption|题注|图题|表题", name, re.I):
                caption_candidates.append(sid)
        if caption_candidates:
            semantic_roles["caption"] = {"style_ids": caption_candidates, "source": "style-name"}

        table_style = table_usage.most_common(1)[0][0] if table_usage else None
        if table_style and table_style != "(direct)":
            semantic_roles["table"] = {"style_id": table_style, "source": "template-usage"}
        table_text_style = table_paragraph_usage.most_common(1)[0][0] if table_paragraph_usage else None
        if table_text_style:
            role: dict[str, Any] = {"source": "template-usage"}
            if table_text_style != "(direct)":
                role["style_id"] = table_text_style
            elif direct_inside:
                role["direct_format"] = json.loads(direct_inside.most_common(1)[0][0])
            semantic_roles["table_text"] = role

        default_paragraph_style = next((sid for sid, item in styles.items() if item.get("type") == "paragraph" and item.get("default")), None)
        normal_style = next((sid for sid, item in styles.items() if item.get("type") == "paragraph" and str(item.get("name", "")).lower() in {"normal", "正文"}), None)
        body_style = default_paragraph_style or normal_style
        if body_style:
            semantic_roles["body"] = {"source": "default-paragraph-style", "style_id": body_style}
        elif paragraph_usage:
            used = paragraph_usage.most_common(1)[0][0]
            role = {"source": "template-usage"}
            if used != "(direct)":
                role["style_id"] = used
            elif direct_outside:
                role["direct_format"] = json.loads(direct_outside.most_common(1)[0][0])
            semantic_roles["body"] = role

        # Preserve the first visible cover/title paragraph's actual direct formatting as a candidate.
        if body is not None:
            first_visible = next((p for p in body.findall("w:p", NS) if _text(p)), None)
            if first_visible is not None:
                ps = first_visible.find("w:pPr/w:pStyle", NS)
                sid = _val(ps)
                title_role = {"source": "first-visible-paragraph"}
                if sid:
                    title_role["style_id"] = sid
                else:
                    title_role["direct_format"] = _paragraph_direct_signature(first_visible)
                semantic_roles["cover_title"] = title_role

        headers: list[dict[str, Any]] = []
        footers: list[dict[str, Any]] = []
        for name in sorted(names):
            if re.fullmatch(r"word/header\d+\.xml", name):
                root = ET.fromstring(z.read(name))
                headers.append({"part": name, "text": " ".join(_text(p) for p in root.findall(".//w:p", NS) if _text(p)), "styles": sorted({_val(p.find("w:pPr/w:pStyle", NS)) or "(direct)" for p in root.findall(".//w:p", NS) if _text(p)})})
            elif re.fullmatch(r"word/footer\d+\.xml", name):
                root = ET.fromstring(z.read(name))
                footers.append({"part": name, "text": " ".join(_text(p) for p in root.findall(".//w:p", NS) if _text(p)), "styles": sorted({_val(p.find("w:pPr/w:pStyle", NS)) or "(direct)" for p in root.findall(".//w:p", NS) if _text(p)})})

        package_hashes = {}
        for name in ("word/styles.xml", "word/numbering.xml", "word/theme/theme1.xml", "word/settings.xml", "word/fontTable.xml"):
            if name in names:
                package_hashes[name] = _canonical_xml_digest(z.read(name))

        return {
            "profile_version": 2,
            "source_file": path.name,
            "source_sha256": "sha256:" + hashlib.sha256(path.read_bytes()).hexdigest(),
            "package_hashes": package_hashes,
            "doc_defaults": doc_defaults,
            "styles": styles,
            "style_name_to_id": name_to_style,
            "semantic_roles": semantic_roles,
            "usage": {
                "paragraph_styles": dict(paragraph_usage),
                "table_paragraph_styles": dict(table_paragraph_usage),
                "table_styles": dict(table_usage),
            },
            "section_geometries": section_geometries,
            "sections": sections,
            "table_contract": table_contract,
            "headers": headers,
            "footers": footers,
        }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("template")
    ap.add_argument("--out", required=True)
    args = ap.parse_args()
    profile = extract_profile(Path(args.template))
    Path(args.out).write_text(json.dumps(profile, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"wrote {args.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
