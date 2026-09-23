#!/usr/bin/env python3
"""Apply deterministic table-alignment and technical-bid page-orientation policy to DOCX."""
from __future__ import annotations

import argparse
import copy
import json
import re
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
qn = lambda name: f"{{{W}}}{name}"


def text_of(node: ET.Element) -> str:
    return "".join(t.text or "" for t in node.findall(".//w:t", NS)).strip()


def orientation(sect: ET.Element) -> str:
    size = sect.find("w:pgSz", NS)
    if size is None:
        return "portrait"
    orient = size.get(qn("orient"))
    try:
        width = int(size.get(qn("w")) or "0")
        height = int(size.get(qn("h")) or "0")
    except ValueError:
        width = height = 0
    return "landscape" if orient == "landscape" or (width > 0 and height > 0 and width > height) else "portrait"


def set_portrait(sect: ET.Element) -> bool:
    size = sect.find("w:pgSz", NS)
    if size is None:
        size = ET.SubElement(sect, qn("pgSz"))
        size.set(qn("w"), "11906")
        size.set(qn("h"), "16838")
        return True
    try:
        width = int(size.get(qn("w")) or "0")
        height = int(size.get(qn("h")) or "0")
    except ValueError:
        width = height = 0
    changed = False
    if width > height > 0:
        size.set(qn("w"), str(height))
        size.set(qn("h"), str(width))
        changed = True
    if qn("orient") in size.attrib:
        del size.attrib[qn("orient")]
        changed = True
    return changed


def section_records(body: ET.Element):
    records = []
    current = []

    def emit(sect):
        nonlocal current
        all_text = " ".join(text_of(el) for el in current if text_of(el)).strip()
        table_paras = {id(p) for el in current for tbl in el.findall(".//w:tbl", NS) for p in tbl.findall(".//w:p", NS)}
        outside = []
        for el in current:
            paras = el.findall(".//w:p", NS) if el.tag != qn("p") else [el]
            for p in paras:
                if id(p) in table_paras:
                    continue
                t = text_of(p)
                if t:
                    outside.append(t)
        contains = "技术偏离表" in all_text.replace(" ", "")
        major = [t for t in outside if re.match(r"^\d+(?:\.\d+)*\s+\S", t)]
        mixed = contains and any("技术偏离表" not in t for t in major)
        records.append({
            "index": len(records) + 1,
            "elements": list(current),
            "sect": sect,
            "orientation": orientation(sect),
            "contains_technical_deviation": contains,
            "mixed_after_deviation": mixed,
            "text_sample": all_text[:240],
        })
        current = []

    for child in list(body):
        if child.tag == qn("sectPr"):
            emit(child)
            continue
        current.append(child)
        if child.tag == qn("p"):
            sect = child.find("w:pPr/w:sectPr", NS)
            if sect is not None:
                emit(sect)
    return records


def repair_document_xml(xml: bytes, technical_bid_policy: bool = False,
                        template: Path | None = None, template_style_json: Path | None = None):
    from template_layout_contract import repair_xml
    active = template or Path(__file__).resolve().parent.parent / "assets" / "default-template.docx"
    updated, report = repair_xml(xml, active, template_style_json)
    # The old keyword is accepted for compatibility, not used to infer layout
    # from business words such as “技术偏离表”. Page geometry is audited separately.
    report.update(centered_cells=0, portrait_sections=[], blocked_mixed_sections=[],
                  template=str(active), legacy_flag_ignored=bool(technical_bid_policy))
    return updated, report


def write_docx(source: Path, output: Path, document_xml: bytes):
    import os
    if source.resolve() == output.resolve():
        raise ValueError("输出不能覆盖原始文档。")
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(dir=output.parent) as work:
        candidate = Path(work) / "layout.docx"
        with zipfile.ZipFile(source, "r") as src, zipfile.ZipFile(candidate, "w") as dst:
            for info in src.infolist():
                data = document_xml if info.filename == "word/document.xml" else src.read(info.filename)
                dst.writestr(info, data)
        os.replace(candidate, output)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("docx", type=Path)
    ap.add_argument("--output", type=Path, required=True)
    ap.add_argument("--technical-bid-policy", action="store_true")
    ap.add_argument("--json-out", type=Path)
    ap.add_argument("--template", type=Path)
    ap.add_argument("--template-style-json", type=Path)
    args = ap.parse_args()

    with zipfile.ZipFile(args.docx, "r") as z:
        xml = z.read("word/document.xml")
    updated, report = repair_document_xml(xml, args.technical_bid_policy, args.template, args.template_style_json)
    if report["blocked_mixed_sections"]:
        report["status"] = "blocked"
        report["message"] = "技术偏离表横向节混入后续普通章节；需先在偏离表结束处建立纵向分节。"
    else:
        report["status"] = "repaired"
        write_docx(args.docx, args.output, updated)
    encoded = json.dumps(report, ensure_ascii=False, indent=2)
    if args.json_out:
        args.json_out.write_text(encoded + "\n", encoding="utf-8")
    print(encoded)
    return 2 if report["status"] == "blocked" else 0


if __name__ == "__main__":
    raise SystemExit(main())
