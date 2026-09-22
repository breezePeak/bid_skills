#!/usr/bin/env python3
"""Deterministic DOCX structural/layout audit for document-layout-review.

Uses only the Python standard library. It does not modify the source file.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET
from punctuation_context import audit_fragment

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


def qn(name: str) -> str:
    return f"{{{W}}}{name}"


HAN = r"\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff"
CJK_PUNCT_CLOSE = "，。；：！？、）】》」』"
CJK_PUNCT_OPEN = "（【《「『"


@dataclass
class Issue:
    code: str
    severity: str
    location: str
    message: str
    evidence: str | None = None
    suggestion: str | None = None


def paragraph_text(p: ET.Element) -> str:
    parts: list[str] = []
    for node in p.iter():
        if node.tag == qn("t"):
            parts.append(node.text or "")
        elif node.tag == qn("tab"):
            parts.append("\t")
        elif node.tag == qn("br"):
            parts.append("\n")
    return "".join(parts)


def short(value: str, limit: int = 96) -> str:
    value = value.replace("\n", "\\n").replace("\t", "\\t")
    return value if len(value) <= limit else value[: limit - 1] + "…"


def audit_text(text: str, location: str) -> list[Issue]:
    issues: list[Issue] = []
    if not text:
        return issues
    if re.search(r"^[ \u3000]+", text):
        issues.append(Issue("leading-space", "warning", location, "段落/单元格以无意义空格开头。", short(text), "删除普通正文中的段首空格；缩进应使用段落格式。"))
    if re.search(r"[ \u3000]+$", text):
        issues.append(Issue("trailing-space", "warning", location, "段落/单元格以无意义空格结尾。", short(text), "删除普通正文中的段尾空格。"))
    if re.search(r" {2,}", text):
        issues.append(Issue("repeated-ascii-space", "warning", location, "存在连续多个普通空格。", short(text), "确认不是英文、代码或对齐用途后归一化。"))
    if re.search(r"\u3000{2,}", text):
        issues.append(Issue("repeated-fullwidth-space", "warning", location, "存在连续多个全角空格。", short(text), "确认不是版式占位后删除。"))
    if re.search(fr"(?<=[{HAN}])[ \u3000]+(?=[{HAN}])", text):
        issues.append(Issue("space-between-han", "warning", location, "相邻中文字符之间存在空格。", short(text), "普通中文正文中删除该空格。"))
    if re.search(fr"[ \u3000]+(?=[{re.escape(CJK_PUNCT_CLOSE)}])", text):
        issues.append(Issue("space-before-cjk-punctuation", "warning", location, "中文闭合/句读标点前存在空格。", short(text), "普通中文正文中删除标点前空格。"))
    if re.search(fr"(?<=[{re.escape(CJK_PUNCT_OPEN)}])[ \u3000]+", text):
        issues.append(Issue("space-after-cjk-opening-punctuation", "warning", location, "中文起始标点后存在空格。", short(text), "普通中文正文中删除起始标点后的空格。"))
    for item in audit_fragment(text):
        issues.append(Issue(
            item.code,
            item.severity,
            location,
            item.message,
            short(text),
            f"建议：{item.expected}" if item.expected else "按中文／英文／代码／公式的局部语境修正标点。",
        ))
    return issues


def iter_body_children(body: ET.Element) -> Iterable[tuple[int, ET.Element]]:
    for idx, child in enumerate(list(body), 1):
        yield idx, child


def _orientation(sect: ET.Element) -> str:
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


def _section_content_width(sect: ET.Element) -> int | None:
    size = sect.find("w:pgSz", NS)
    mar = sect.find("w:pgMar", NS)
    if size is None or mar is None:
        return None
    try:
        width = int(size.get(qn("w")) or "0")
        left = int(mar.get(qn("left")) or "0")
        right = int(mar.get(qn("right")) or "0")
    except ValueError:
        return None
    return width - left - right if width > 0 else None


def _section_records(body: ET.Element) -> list[dict]:
    records: list[dict] = []
    current: list[ET.Element] = []

    def emit(sect: ET.Element) -> None:
        nonlocal current
        text = " ".join(paragraph_text(p) for el in current for p in (el.findall(".//w:p", NS) if el.tag != qn("p") else [el]) if paragraph_text(p)).strip()
        table_paras = {id(p) for el in current for tbl in el.findall(".//w:tbl", NS) for p in tbl.findall(".//w:p", NS)}
        outside = []
        for el in current:
            paras = el.findall(".//w:p", NS) if el.tag != qn("p") else [el]
            for p in paras:
                if id(p) not in table_paras:
                    t = paragraph_text(p).strip()
                    if t:
                        outside.append(t)
        contains_deviation = "技术偏离表" in text.replace(" ", "")
        major_headings = [t for t in outside if re.match(r"^\d+(?:\.\d+)*\s+\S", t)]
        mixed = contains_deviation and any("技术偏离表" not in t for t in major_headings)
        records.append({
            "index": len(records) + 1,
            "orientation": _orientation(sect),
            "contains_technical_deviation": contains_deviation,
            "mixed_after_deviation": mixed,
            "content_width_twips": _section_content_width(sect),
            "table_element_ids": [id(tbl) for el in current for tbl in el.findall(".//w:tbl", NS)],
            "text_sample": text[:240],
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


def section_geometry(root: ET.Element) -> dict[str, int | None]:
    sect = root.find(".//w:sectPr", NS)
    if sect is None:
        return {"page_width_twips": None, "page_height_twips": None, "content_width_twips": None}
    pgsz = sect.find("w:pgSz", NS)
    pgmar = sect.find("w:pgMar", NS)
    def val(el: ET.Element | None, attr: str) -> int | None:
        if el is None:
            return None
        raw = el.get(qn(attr))
        try:
            return int(raw) if raw is not None else None
        except ValueError:
            return None
    width = val(pgsz, "w")
    height = val(pgsz, "h")
    left = val(pgmar, "left")
    right = val(pgmar, "right")
    content = width - left - right if None not in (width, left, right) else None
    return {"page_width_twips": width, "page_height_twips": height, "content_width_twips": content}


def audit_docx(path: Path, technical_bid_policy: bool = False) -> dict:
    with zipfile.ZipFile(path) as zf:
        try:
            xml = zf.read("word/document.xml")
        except KeyError as exc:
            raise SystemExit("DOCX 缺少 word/document.xml") from exc
    root = ET.fromstring(xml)
    body = root.find("w:body", NS)
    if body is None:
        raise SystemExit("DOCX document.xml 缺少 w:body")

    sections = _section_records(body)
    table_content_width = {table_id: section.get("content_width_twips") for section in sections for table_id in section.get("table_element_ids", [])}
    enforce_technical_bid = technical_bid_policy or any(s.get("contains_technical_deviation") for s in sections)

    issues: list[Issue] = []
    if enforce_technical_bid:
        landscape_sections = []
        for section in sections:
            if section.get("orientation") != "landscape":
                continue
            landscape_sections.append(section.get("index"))
            if not section.get("contains_technical_deviation"):
                issues.append(Issue(
                    "technical-bid-landscape-outside-deviation-table",
                    "error",
                    f"section:{section.get('index')}",
                    "技术标只有“技术偏离表”所在节允许横向，其他内容必须纵向。",
                    section.get("text_sample"),
                    "将该节恢复为纵向；宽表应通过列宽、换行或分页优化解决。",
                ))
            elif section.get("mixed_after_deviation"):
                issues.append(Issue(
                    "technical-bid-landscape-mixed-content",
                    "error",
                    f"section:{section.get('index')}",
                    "横向技术偏离表节包含后续普通章节内容；普通章节不得继承横向方向。",
                    section.get("text_sample"),
                    "在技术偏离表结束后建立新的纵向节，再放置后续正文。",
                ))
        if len(landscape_sections) > 1:
            issues.append(Issue(
                "technical-bid-multiple-landscape-sections",
                "error",
                "document",
                "技术标存在多个横向节；最多只允许技术偏离表所在节横向。",
                str(landscape_sections),
                "将其他横向节恢复为纵向。",
            ))
    para_count = 0
    table_count = 0

    # Paragraph text / hard-break audit, including table-cell paragraphs.
    for para_count, p in enumerate(root.findall(".//w:p", NS), 1):
        text = paragraph_text(p)
        location = f"paragraph:{para_count}"
        issues.extend(audit_text(text, location))
        if p.find(".//w:br", NS) is not None:
            issues.append(Issue("hard-line-break", "info", location, "段落中包含硬换行。", short(text), "确认该换行确有排版目的；普通正文优先拆成独立段落。"))

    # Consecutive top-level empty paragraphs.
    empty_run = 0
    empty_start = None
    body_para_index = 0
    for _, child in iter_body_children(body):
        if child.tag != qn("p"):
            if empty_run >= 2:
                issues.append(Issue("consecutive-empty-paragraphs", "warning", f"body-paragraph:{empty_start}", f"存在连续 {empty_run} 个空段落。", suggestion="若不是刻意留白，保留最多一个。"))
            empty_run = 0
            empty_start = None
            continue
        body_para_index += 1
        if paragraph_text(child).strip() == "":
            if empty_run == 0:
                empty_start = body_para_index
            empty_run += 1
        else:
            if empty_run >= 2:
                issues.append(Issue("consecutive-empty-paragraphs", "warning", f"body-paragraph:{empty_start}", f"存在连续 {empty_run} 个空段落。", suggestion="若不是刻意留白，保留最多一个。"))
            empty_run = 0
            empty_start = None
    if empty_run >= 2:
        issues.append(Issue("consecutive-empty-paragraphs", "warning", f"body-paragraph:{empty_start}", f"存在连续 {empty_run} 个空段落。", suggestion="若不是刻意留白，保留最多一个。"))

    geometry = section_geometry(root)
    content_width = geometry["content_width_twips"]

    for table_count, tbl in enumerate(root.findall(".//w:tbl", NS), 1):
        loc = f"table:{table_count}"
        tblw = tbl.find("w:tblPr/w:tblW", NS)
        if tblw is None:
            issues.append(Issue("table-width-missing", "warning", loc, "表格缺少显式总宽度。", suggestion="为最终 DOCX 设置确定的表格宽度。"))
        else:
            typ = tblw.get(qn("type"))
            width_raw = tblw.get(qn("w"))
            if typ == "pct":
                issues.append(Issue("table-percent-width", "info", loc, "表格使用百分比宽度。", f"w:tblW={width_raw}, type=pct", "布局敏感文档可考虑固定最终宽度并同步单元格宽度。"))

        grid = tbl.findall("w:tblGrid/w:gridCol", NS)
        widths: list[int] = []
        for col in grid:
            raw = col.get(qn("w"))
            try:
                widths.append(int(raw or "0"))
            except ValueError:
                widths.append(0)
        if not widths or any(w <= 0 for w in widths):
            issues.append(Issue("table-grid-width-invalid", "warning", loc, "表格列网格缺少有效固定宽度。", suggestion="为每一列计算并写入稳定列宽。"))
        table_width_limit = table_content_width.get(id(tbl)) or content_width
        if widths and table_width_limit is not None and sum(widths) > table_width_limit * 1.02:
            issues.append(Issue("table-grid-overflow", "error", loc, "表格列宽总和超过所在节可用页面宽度。", f"grid={sum(widths)} twips, content={table_width_limit} twips", "重新分配列宽；技术标普通表格不得通过改横向页面解决。"))

        rows = tbl.findall("w:tr", NS)
        if rows:
            trpr = rows[0].find("w:trPr", NS)
            repeat = trpr.find("w:tblHeader", NS) if trpr is not None else None
            if repeat is None:
                issues.append(Issue("table-header-not-repeating", "info", loc, "首行未标记为跨页重复表头。", suggestion="长表格在渲染器支持时启用重复表头。"))

        cells = tbl.findall(".//w:tc", NS)
        for cell_index, tc in enumerate(cells, 1):
            c_loc = f"{loc}/cell:{cell_index}"
            tcpr = tc.find("w:tcPr", NS)
            valign = tcpr.find("w:vAlign", NS) if tcpr is not None else None
            value = valign.get(qn("val")) if valign is not None else None
            if value != "center":
                issues.append(Issue("table-cell-not-vertically-centered", "error", c_loc, "表格单元格必须上下居中。", f"vAlign={value or 'unset'}", "在 w:tcPr 中设置 w:vAlign=\"center\"；这是硬规则，不能只依赖表格样式。"))
            cell_text = "\n".join(paragraph_text(p) for p in tc.findall("w:p", NS))
            issues.extend(audit_text(cell_text, c_loc))

    sev_counts = {"error": 0, "warning": 0, "info": 0}
    for issue in issues:
        sev_counts[issue.severity] = sev_counts.get(issue.severity, 0) + 1

    return {
        "file": str(path),
        "summary": {
            "paragraphs": para_count,
            "tables": table_count,
            "issues": len(issues),
            "severity": sev_counts,
            "sections": [{k: v for k, v in section.items() if k != "table_element_ids"} for section in sections],
            **geometry,
        },
        "issues": [asdict(i) for i in issues],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit DOCX layout/whitespace structure without modifying it.")
    parser.add_argument("docx", type=Path)
    parser.add_argument("--json-out", type=Path)
    parser.add_argument("--technical-bid-policy", action="store_true", help="Enforce portrait except the technical-deviation-table section.")
    args = parser.parse_args()

    if not args.docx.is_file():
        parser.error(f"file not found: {args.docx}")
    try:
        result = audit_docx(args.docx, technical_bid_policy=args.technical_bid_policy)
    except zipfile.BadZipFile:
        parser.error("input is not a valid DOCX/ZIP file")
    encoded = json.dumps(result, ensure_ascii=False, indent=2)
    if args.json_out:
        args.json_out.parent.mkdir(parents=True, exist_ok=True)
        args.json_out.write_text(encoded + "\n", encoding="utf-8")
    print(encoded)
    return 2 if result["summary"]["severity"]["error"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
