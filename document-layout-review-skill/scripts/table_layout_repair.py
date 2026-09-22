#!/usr/bin/env python3
"""Deterministic DOCX table reflow for document-layout-review.

Goals:
- keep the active template's fonts/borders/fills/styles;
- repair structural layout only (widths, alignment, margins, header repeat);
- use content-aware widths instead of equal-width or stale template widths;
- use a stable semantic width contract for the technical deviation table;
- never rotate ordinary technical-bid tables to landscape.
"""
from __future__ import annotations

import argparse
import json
import math
import re
import zipfile
from collections import defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}

def qn(name: str) -> str:
    return f"{{{W}}}{name}"

ZERO_WIDTH = {"\u200b", "\u200c", "\u200d", "\ufeff"}
TECH_HEADERS = ["序号", "标的名称", "招标技术要求", "投标响应内容", "偏离程度", "备注"]
COMPACT_KEYS = ("序号", "编号", "序", "状态", "偏离程度", "是否", "等级", "日期", "时间", "数量", "单位", "版本")
MEDIUM_KEYS = ("名称", "对象", "阶段", "类别", "主要输入", "主要输出", "甲方接口", "成果", "工作对象")
NARRATIVE_KEYS = ("说明", "内容", "要求", "措施", "重点", "影响", "思路", "动作", "产出", "边界", "响应", "备注", "依据", "处理", "工作", "方法", "放行", "审查")


def clean_text(value: str) -> str:
    return "".join(ch for ch in value if ch not in ZERO_WIDTH).strip()


def node_text(node: ET.Element) -> str:
    return clean_text("".join(t.text or "" for t in node.findall(".//w:t", NS)))


def ensure_child(parent: ET.Element, tag: str, before: tuple[str, ...] = ()) -> ET.Element:
    child = parent.find(f"w:{tag}", NS)
    if child is not None:
        return child
    child = ET.Element(qn(tag))
    if before:
        for i, existing in enumerate(list(parent)):
            if existing.tag in {qn(x) for x in before}:
                parent.insert(i, child)
                return child
    parent.append(child)
    return child


def set_val(parent: ET.Element, tag: str, value: str, attr: str = "val") -> ET.Element:
    child = ensure_child(parent, tag)
    child.set(qn(attr), value)
    return child


def visual_width(text: str) -> float:
    total = 0.0
    for ch in clean_text(text):
        o = ord(ch)
        if 0x3400 <= o <= 0x9FFF or 0xF900 <= o <= 0xFAFF:
            total += 1.0
        elif ch.isspace():
            total += 0.3
        else:
            total += 0.55
    return total


def cell_span(tc: ET.Element) -> int:
    span = tc.find("w:tcPr/w:gridSpan", NS)
    if span is None:
        return 1
    try:
        return max(1, int(span.get(qn("val")) or "1"))
    except ValueError:
        return 1


def logical_column_count(tbl: ET.Element) -> int:
    grid = tbl.findall("w:tblGrid/w:gridCol", NS)
    if grid:
        return len(grid)
    count = 0
    for row in tbl.findall("w:tr", NS):
        count = max(count, sum(cell_span(tc) for tc in row.findall("w:tc", NS)))
    return count


def row_cells_by_column(row: ET.Element, count: int):
    result = []
    col = 0
    for tc in row.findall("w:tc", NS):
        span = cell_span(tc)
        result.append((tc, col, min(count, col + span)))
        col += span
    return result


def headers(tbl: ET.Element, count: int) -> list[str]:
    rows = tbl.findall("w:tr", NS)
    if not rows:
        return [""] * count
    out = [""] * count
    for tc, start, end in row_cells_by_column(rows[0], count):
        value = node_text(tc)
        for i in range(start, end):
            out[i] = value
    return out


def is_technical_deviation_table(tbl: ET.Element) -> bool:
    count = logical_column_count(tbl)
    hs = [re.sub(r"\s+", "", h) for h in headers(tbl, count)]
    joined = "|".join(hs)
    return all(h in joined for h in TECH_HEADERS)


def normalize_technical_deviation_structure(tbl: ET.Element) -> int:
    """保证默认技术偏离表使用已确认的 7 个物理单元格。

    表头仍显示 6 个分组；“招标技术要求”跨两个物理列：
    第 3 列为星号/三角标识，第 4 列为技术要求正文。
    已经是 7 列时绝不压成 6 列；历史文档若曾被旧版误压成 6 列，则恢复空的标识列。
    """
    if not is_technical_deviation_table(tbl):
        return 0
    rows = tbl.findall("w:tr", NS)
    if not rows:
        return 0
    grid = tbl.find("w:tblGrid", NS)
    grid_count = len(grid.findall("w:gridCol", NS)) if grid is not None else 0
    if grid_count == 7:
        return 0
    if grid_count != 6:
        return 0
    header_cells = rows[0].findall("w:tc", NS)
    if len(header_cells) != 6:
        return 0
    # 让“招标技术要求”表头跨第 3、4 两个物理列。
    tcpr = header_cells[2].find("w:tcPr", NS)
    if tcpr is None:
        tcpr = ET.Element(qn("tcPr")); header_cells[2].insert(0, tcpr)
    span = tcpr.find("w:gridSpan", NS)
    if span is None:
        span = ET.SubElement(tcpr, qn("gridSpan"))
    span.set(qn("val"), "2")
    # 每个正文行在原第 3 列前插入一个空的“星号/三角标识”单元格。
    for row in rows[1:]:
        cells = row.findall("w:tc", NS)
        if len(cells) != 6:
            continue
        marker = ET.Element(qn("tc"))
        marker_pr = ET.SubElement(marker, qn("tcPr"))
        ET.SubElement(marker_pr, qn("vAlign"), {qn("val"): "center"})
        para = ET.SubElement(marker, qn("p"))
        row.insert(list(row).index(cells[2]), marker)
    if grid is None:
        grid = ET.Element(qn("tblGrid")); tbl.insert(1, grid)
    for col in list(grid.findall("w:gridCol", NS)):
        grid.remove(col)
    for _ in range(7):
        ET.SubElement(grid, qn("gridCol"), {qn("w"): "1"})
    return 1 + sum(1 for row in rows[1:] if len(row.findall("w:tc", NS)) == 7)

def role_for_header(text: str) -> str:
    t = re.sub(r"\s+", "", text)
    if any(k in t for k in COMPACT_KEYS):
        return "compact"
    if any(k in t for k in NARRATIVE_KEYS):
        return "narrative"
    if any(k in t for k in MEDIUM_KEYS):
        return "medium"
    return "normal"


def pressure_by_column(tbl: ET.Element, count: int) -> tuple[list[float], list[float], list[str]]:
    vals: list[list[float]] = [[] for _ in range(count)]
    hs = headers(tbl, count)
    for r_index, row in enumerate(tbl.findall("w:tr", NS)):
        for tc, start, end in row_cells_by_column(row, count):
            width = visual_width(node_text(tc))
            span = max(1, end - start)
            contribution = width / span
            for col in range(start, end):
                vals[col].append(contribution * (0.65 if r_index == 0 else 1.0))
    avg = [sum(v) / len(v) if v else 0.0 for v in vals]
    maxv = [max(v) if v else 0.0 for v in vals]
    return avg, maxv, hs


def _bounds(role: str, count: int) -> tuple[float, float, float]:
    # (minimum fraction, maximum fraction, role multiplier)
    if count <= 2:
        base_min = 0.22
    elif count == 3:
        base_min = 0.14
    elif count == 4:
        base_min = 0.10
    else:
        base_min = 0.06
    if role == "compact":
        return (max(0.05, base_min * 0.75), 0.14 if count >= 4 else 0.22, 0.65)
    if role == "medium":
        return (max(0.10, base_min), 0.26, 1.0)
    if role == "narrative":
        return (max(0.13, base_min), 0.46, 1.35)
    return (max(0.09, base_min), 0.34, 1.0)


def allocate_fractions(weights: list[float], mins: list[float], maxs: list[float]) -> list[float]:
    n = len(weights)
    if not n:
        return []
    # Ensure minimums are feasible.
    smin = sum(mins)
    if smin >= 0.98:
        mins = [m * 0.92 / smin for m in mins]
    widths = mins[:]
    remaining = 1.0 - sum(widths)
    active = set(range(n))
    weights = [max(0.01, w) for w in weights]
    while remaining > 1e-9 and active:
        total_w = sum(weights[i] for i in active)
        consumed = 0.0
        saturated = []
        for i in list(active):
            share = remaining * weights[i] / total_w
            room = maxs[i] - widths[i]
            add = min(share, max(0.0, room))
            widths[i] += add
            consumed += add
            if room <= share + 1e-9:
                saturated.append(i)
        remaining -= consumed
        for i in saturated:
            active.discard(i)
        if consumed < 1e-10:
            break
    if remaining > 1e-8:
        # Distribute any residual evenly where possible.
        for i in range(n):
            room = maxs[i] - widths[i]
            add = min(room, remaining / max(1, n - i))
            widths[i] += add
            remaining -= add
    total = sum(widths)
    return [w / total for w in widths]


def infer_fractions(tbl: ET.Element, count: int, template_ratios: list[float] | None = None) -> tuple[list[float], list[str]]:
    hs = headers(tbl, count)
    roles = [role_for_header(h) for h in hs]
    if count == 7 and is_technical_deviation_table(tbl):
        roles = ["compact", "medium", "compact", "narrative", "narrative", "compact", "narrative"]
    if template_ratios is not None and len(template_ratios) == count and abs(sum(template_ratios) - 1.0) < 0.02:
        return template_ratios, roles
    if count == 7 and is_technical_deviation_table(tbl):
        # 7 个物理列：序号、标的名称、技术要求标识、技术要求正文、投标响应、偏离程度、备注。
        # 正常情况下优先复用已确认模板的真实 7 列比例；这里只是无模板时的兜底。
        return [0.0574, 0.0898, 0.0400, 0.3502, 0.3094, 0.1002, 0.0530], ["compact", "medium", "compact", "narrative", "narrative", "compact", "narrative"]
    if count == 6 and is_technical_deviation_table(tbl):
        # 兼容用户自定义的六列技术偏离表，不强制扩成七列。
        return [0.05, 0.12, 0.23, 0.27, 0.10, 0.23], ["compact", "medium", "narrative", "narrative", "compact", "narrative"]
    avg, maxv, hs = pressure_by_column(tbl, count)
    weights, mins, maxs = [], [], []
    for i, role in enumerate(roles):
        mn, mx, mult = _bounds(role, count)
        signal = 1.0 + math.log1p(avg[i] * 0.55 + maxv[i] * 0.45)
        weights.append(signal * mult)
        mins.append(mn)
        maxs.append(mx)
    return allocate_fractions(weights, mins, maxs), roles


def section_content_width(sect: ET.Element | None) -> int:
    if sect is None:
        return 8306
    size = sect.find("w:pgSz", NS)
    mar = sect.find("w:pgMar", NS)
    try:
        width = int(size.get(qn("w"))) if size is not None and size.get(qn("w")) else 11906
        left = int(mar.get(qn("left"))) if mar is not None and mar.get(qn("left")) else 1800
        right = int(mar.get(qn("right"))) if mar is not None and mar.get(qn("right")) else 1800
    except ValueError:
        return 8306
    return max(1000, width - left - right)


def table_section_widths(body: ET.Element) -> dict[int, int]:
    result: dict[int, int] = {}
    current: list[ET.Element] = []
    def emit(sect: ET.Element):
        width = section_content_width(sect)
        for el in current:
            for tbl in el.findall(".//w:tbl", NS):
                result[id(tbl)] = width
        current.clear()
    for child in list(body):
        if child.tag == qn("sectPr"):
            emit(child)
            continue
        current.append(child)
        if child.tag == qn("p"):
            sect = child.find("w:pPr/w:sectPr", NS)
            if sect is not None:
                emit(sect)
    return result


def set_table_margins(tblpr: ET.Element, top=70, bottom=70, left=90, right=90):
    mar = ensure_child(tblpr, "tblCellMar")
    for side, value in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        node = mar.find(f"w:{side}", NS)
        if node is None:
            node = ET.SubElement(mar, qn(side))
        node.set(qn("w"), str(value))
        node.set(qn("type"), "dxa")


def set_paragraph_alignment(p: ET.Element, value: str):
    ppr = p.find("w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(qn("pPr"))
        p.insert(0, ppr)
    set_val(ppr, "jc", value)
    ind = ppr.find("w:ind", NS)
    if ind is None:
        ind = ET.SubElement(ppr, qn("ind"))
    ind.set(qn("firstLine"), "0")
    ind.set(qn("firstLineChars"), "0")
    spacing = ppr.find("w:spacing", NS)
    if spacing is None:
        spacing = ET.SubElement(ppr, qn("spacing"))
    spacing.set(qn("before"), "0")
    spacing.set(qn("after"), "0")
    spacing.set(qn("line"), "240")
    spacing.set(qn("lineRule"), "auto")


def repair_table(tbl: ET.Element, content_width: int, template_ratios: list[float] | None = None, template_table_style: str | None = None, template_text_style: str | None = None) -> dict:
    structure_changes = normalize_technical_deviation_structure(tbl)
    count = logical_column_count(tbl)
    if count <= 0:
        return {"changed": False, "reason": "no-columns"}
    fractions, roles = infer_fractions(tbl, count, template_ratios)
    total_width = max(1000, int(content_width * 0.995))
    widths = [max(1, round(total_width * f)) for f in fractions]
    widths[-1] += total_width - sum(widths)

    tblpr = tbl.find("w:tblPr", NS)
    if tblpr is None:
        tblpr = ET.Element(qn("tblPr"))
        tbl.insert(0, tblpr)
    if template_table_style:
        style = ensure_child(tblpr, "tblStyle")
        style.set(qn("val"), template_table_style)
    tblw = ensure_child(tblpr, "tblW")
    tblw.set(qn("w"), str(total_width))
    tblw.set(qn("type"), "dxa")
    layout = ensure_child(tblpr, "tblLayout")
    layout.set(qn("type"), "fixed")
    jc = ensure_child(tblpr, "jc")
    jc.set(qn("val"), "center")
    set_table_margins(tblpr)

    grid = tbl.find("w:tblGrid", NS)
    if grid is None:
        grid = ET.Element(qn("tblGrid"))
        tbl.insert(1 if tbl.find("w:tblPr", NS) is not None else 0, grid)
    for child in list(grid):
        if child.tag == qn("gridCol"):
            grid.remove(child)
    for w in widths:
        ET.SubElement(grid, qn("gridCol"), {qn("w"): str(w)})

    rows = tbl.findall("w:tr", NS)
    for r_index, row in enumerate(rows):
        trpr = row.find("w:trPr", NS)
        if trpr is None:
            trpr = ET.Element(qn("trPr"))
            row.insert(0, trpr)
        if r_index == 0:
            set_val(trpr, "tblHeader", "1")
        for tc, start, end in row_cells_by_column(row, count):
            tcpr = tc.find("w:tcPr", NS)
            if tcpr is None:
                tcpr = ET.Element(qn("tcPr"))
                tc.insert(0, tcpr)
            tcw = ensure_child(tcpr, "tcW")
            tcw.set(qn("w"), str(sum(widths[start:end])))
            tcw.set(qn("type"), "dxa")
            valign = ensure_child(tcpr, "vAlign")
            valign.set(qn("val"), "center")
            # Remove zero-width artifacts from visible cell text without touching run formatting.
            for t in tc.findall(".//w:t", NS):
                if t.text:
                    t.text = "".join(ch for ch in t.text if ch not in ZERO_WIDTH)
            # Headers and compact semantic columns are centered; narrative cells stay left aligned.
            align = "center" if r_index == 0 or (start < len(roles) and roles[start] == "compact") else "left"
            for p in tc.findall("w:p", NS):
                if template_text_style:
                    ppr = p.find("w:pPr", NS)
                    if ppr is None:
                        ppr = ET.Element(qn("pPr")); p.insert(0, ppr)
                    pstyle = ppr.find("w:pStyle", NS)
                    if pstyle is None:
                        pstyle = ET.Element(qn("pStyle")); ppr.insert(0, pstyle)
                    pstyle.set(qn("val"), template_text_style)
                set_paragraph_alignment(p, align)
    return {
        "changed": True,
        "technical_deviation": is_technical_deviation_table(tbl),
        "columns": count,
        "content_width_twips": content_width,
        "table_width_twips": total_width,
        "column_widths_twips": widths,
        "column_fractions": [round(w / total_width, 4) for w in widths],
        "roles": roles,
        "structure_changes": structure_changes,
    }


def header_signature(tbl: ET.Element) -> tuple[str, ...]:
    count = logical_column_count(tbl)
    return tuple(re.sub(r"\s+", "", h) for h in headers(tbl, count))


def template_contract(template: Path | None):
    if template is None:
        return {}, None, None
    with zipfile.ZipFile(template, "r") as z:
        root = ET.fromstring(z.read("word/document.xml"))
    samples = {}
    table_style_counts = defaultdict(int)
    text_style_counts = defaultdict(int)
    for tbl in root.findall(".//w:tbl", NS):
        count = logical_column_count(tbl)
        grid = tbl.findall("w:tblGrid/w:gridCol", NS)
        vals = []
        for c in grid:
            try: vals.append(int(c.get(qn("w")) or "0"))
            except ValueError: vals.append(0)
        if count and len(vals) == count and sum(vals) > 0:
            samples[header_signature(tbl)] = [v / sum(vals) for v in vals]
        style = tbl.find("w:tblPr/w:tblStyle", NS)
        if style is not None and style.get(qn("val")):
            table_style_counts[style.get(qn("val"))] += 1
        for p in tbl.findall(".//w:p", NS):
            ps = p.find("w:pPr/w:pStyle", NS)
            if ps is not None and ps.get(qn("val")):
                text_style_counts[ps.get(qn("val"))] += 1
    table_style = max(table_style_counts, key=table_style_counts.get) if table_style_counts else None
    text_style = max(text_style_counts, key=text_style_counts.get) if text_style_counts else None
    return samples, table_style, text_style


def repair_document_xml(xml: bytes, template: Path | None = None) -> tuple[bytes, dict]:
    root = ET.fromstring(xml)
    body = root.find("w:body", NS)
    if body is None:
        raise ValueError("DOCX document.xml lacks w:body")
    section_widths = table_section_widths(body)
    samples, template_table_style, template_text_style = template_contract(template)
    tables = root.findall(".//w:tbl", NS)
    reports = []
    default_width = 8306
    for i, tbl in enumerate(tables, 1):
        sig = header_signature(tbl)
        rep = repair_table(tbl, section_widths.get(id(tbl), default_width), samples.get(sig), template_table_style, template_text_style)
        rep["template_widths_used"] = sig in samples
        rep["table"] = i
        reports.append(rep)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), {"tables": reports, "table_count": len(tables)}


def write_docx(source: Path, output: Path, document_xml: bytes):
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source, "r") as src, zipfile.ZipFile(output, "w") as dst:
        for info in src.infolist():
            data = document_xml if info.filename == "word/document.xml" else src.read(info.filename)
            dst.writestr(info, data)


def main() -> int:
    ap = argparse.ArgumentParser(description="Reflow DOCX tables with content-aware fixed widths while preserving template styles.")
    ap.add_argument("docx", type=Path)
    ap.add_argument("--output", type=Path, required=True)
    ap.add_argument("--json-out", type=Path)
    ap.add_argument("--template", type=Path, help="Active DOCX template; matching table widths/styles are reused before content-aware fallback.")
    args = ap.parse_args()
    with zipfile.ZipFile(args.docx, "r") as z:
        xml = z.read("word/document.xml")
    updated, report = repair_document_xml(xml, args.template)
    write_docx(args.docx, args.output, updated)
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.json_out:
        args.json_out.parent.mkdir(parents=True, exist_ok=True)
        args.json_out.write_text(text + "\n", encoding="utf-8")
    print(text)
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
