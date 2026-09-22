#!/usr/bin/env python3
"""Deterministic table-layout audit for document-layout-review.

This script does not decide whether a table is aesthetically perfect. It detects
objective layout risks and produces semantic-review candidates for issues that
require content understanding, especially vertical merging of repeated grouping
cells.

Checks:
- physical table width vs printable section width;
- column widths vs header/body content pressure;
- abrupt / unjustified width imbalance;
- one-character-per-line / cramped-header risk;
- autofit + long unbreakable-token page-stretch risk;
- approved physical-grid contract for technical-deviation tables when supplied
  by the active template style JSON;
- consecutive repeated cells that are plausible vertical-merge candidates.

Exit code:
- 0: no error-severity issue;
- 2: one or more error-severity issues;
- 3: invalid input / parse failure.
"""
from __future__ import annotations

import argparse
import json
import math
import re
import statistics
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}

ZERO_WIDTH = {"\u200b", "\u200c", "\u200d", "\ufeff"}
TECH_HEADERS = ["序号", "标的名称", "招标技术要求", "投标响应内容", "偏离程度", "备注"]

COMPACT_KEYS = (
    "序号", "编号", "序", "状态", "偏离程度", "是否", "等级", "日期", "时间",
    "数量", "单位", "版本", "序列", "代码",
)
MEDIUM_KEYS = (
    "名称", "对象", "阶段", "类别", "模块", "工作包", "责任单位", "责任部门",
    "任务", "成果类型", "标的",
)
NARRATIVE_KEYS = (
    "说明", "内容", "要求", "措施", "重点", "影响", "思路", "动作", "产出",
    "边界", "响应", "备注", "依据", "处理", "工作", "方法", "放行", "审查",
    "目标", "风险", "输入", "输出", "结论",
)
MERGE_FRIENDLY_KEYS = (
    "标的名称", "项目名称", "工作对象", "对象", "类别", "阶段", "模块", "工作包",
    "责任单位", "责任部门", "任务名称", "成果类型", "所属项目", "所属模块",
)
NO_MERGE_KEYS = (
    "序号", "编号", "状态", "是否", "偏离程度", "等级", "日期", "时间", "数量",
    "单位", "版本", "备注", "结果", "结论",
)

ASCII_TOKEN_RE = re.compile(r"[A-Za-z0-9_./:\\@#?&=%+\-]{16,}")
HAN_RE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]")


def qn(name: str) -> str:
    return f"{{{W}}}{name}"


def clean_text(value: str) -> str:
    return "".join(ch for ch in value if ch not in ZERO_WIDTH).strip()


def node_text(node: ET.Element) -> str:
    return clean_text("".join(t.text or "" for t in node.findall(".//w:t", NS)))


def visual_width(text: str) -> float:
    """Approximate text pressure in CJK-em units."""
    total = 0.0
    for ch in clean_text(text):
        o = ord(ch)
        if 0x3400 <= o <= 0x9FFF or 0xF900 <= o <= 0xFAFF:
            total += 1.0
        elif ch.isspace():
            total += 0.3
        elif ch.isupper():
            total += 0.65
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


def normalized_header(text: str) -> str:
    return re.sub(r"\s+", "", text)


def is_technical_deviation_table(tbl: ET.Element) -> bool:
    count = logical_column_count(tbl)
    hs = [normalized_header(h) for h in headers(tbl, count)]
    joined = "|".join(hs)
    return all(h in joined for h in TECH_HEADERS)


def role_for_header(text: str) -> str:
    t = normalized_header(text)
    if any(k in t for k in COMPACT_KEYS):
        return "compact"
    if any(k in t for k in NARRATIVE_KEYS):
        return "narrative"
    if any(k in t for k in MEDIUM_KEYS):
        return "medium"
    return "normal"


def merge_friendly_header(text: str) -> bool:
    t = normalized_header(text)
    if not t or any(k in t for k in NO_MERGE_KEYS):
        return False
    return any(k in t for k in MERGE_FRIENDLY_KEYS)


def percentile(values: list[float], q: float) -> float:
    if not values:
        return 0.0
    vals = sorted(values)
    if len(vals) == 1:
        return vals[0]
    pos = (len(vals) - 1) * q
    lo = int(math.floor(pos))
    hi = int(math.ceil(pos))
    if lo == hi:
        return vals[lo]
    frac = pos - lo
    return vals[lo] * (1 - frac) + vals[hi] * frac


def grid_widths(tbl: ET.Element, count: int) -> list[int]:
    grid = tbl.findall("w:tblGrid/w:gridCol", NS)
    if len(grid) == count:
        vals = []
        for c in grid:
            try:
                vals.append(max(0, int(c.get(qn("w")) or "0")))
            except ValueError:
                vals.append(0)
        if all(v > 0 for v in vals):
            return vals

    # Fall back to first row tcW distributed across spans.
    rows = tbl.findall("w:tr", NS)
    vals = [0] * count
    if rows:
        for tc, start, end in row_cells_by_column(rows[0], count):
            w = tc.find("w:tcPr/w:tcW", NS)
            try:
                value = int(w.get(qn("w")) or "0") if w is not None else 0
            except ValueError:
                value = 0
            span = max(1, end - start)
            each = value // span if value > 0 else 0
            for i in range(start, end):
                vals[i] = max(vals[i], each)
    return vals


def vmerge_state(tc: ET.Element) -> str | None:
    node = tc.find("w:tcPr/w:vMerge", NS)
    if node is None:
        return None
    return node.get(qn("val")) or "continue"


def has_existing_vertical_merge(tbl: ET.Element, col: int, start_row: int, end_row: int) -> bool:
    rows = tbl.findall("w:tr", NS)
    for row_idx in range(start_row, min(end_row + 1, len(rows))):
        for tc, start, end in row_cells_by_column(rows[row_idx], logical_column_count(tbl)):
            if start <= col < end and vmerge_state(tc) is not None:
                return True
    return False


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


def tables_in_element(el: ET.Element) -> list[ET.Element]:
    result = []
    if el.tag == qn("tbl"):
        result.append(el)
    result.extend(el.findall(".//w:tbl", NS))
    # De-duplicate in case nested lookups overlap.
    seen = set()
    out = []
    for t in result:
        if id(t) not in seen:
            seen.add(id(t))
            out.append(t)
    return out


def table_section_widths(body: ET.Element) -> dict[int, int]:
    result: dict[int, int] = {}
    current: list[ET.Element] = []

    def emit(sect: ET.Element):
        width = section_content_width(sect)
        for el in current:
            for tbl in tables_in_element(el):
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

    # Documents normally end with sectPr, but be defensive.
    if current:
        fallback = body.find("w:sectPr", NS)
        width = section_content_width(fallback)
        for el in current:
            for tbl in tables_in_element(el):
                result[id(tbl)] = width
    return result


def _bounds(role: str, count: int) -> tuple[float, float, float]:
    if count <= 2:
        base_min = 0.22
    elif count == 3:
        base_min = 0.14
    elif count == 4:
        base_min = 0.10
    else:
        base_min = 0.055

    if role == "compact":
        return max(0.045, base_min * 0.75), (0.15 if count >= 4 else 0.24), 0.65
    if role == "medium":
        return max(0.085, base_min), 0.28, 1.00
    if role == "narrative":
        return max(0.12, base_min), 0.48, 1.40
    return max(0.075, base_min), 0.36, 1.00


def allocate_fractions(weights: list[float], mins: list[float], maxs: list[float]) -> list[float]:
    n = len(weights)
    if not n:
        return []
    mins = mins[:]
    maxs = maxs[:]
    smin = sum(mins)
    if smin >= 0.96:
        mins = [m * 0.92 / smin for m in mins]
    widths = mins[:]
    remaining = 1.0 - sum(widths)
    active = set(range(n))
    safe_weights = [max(0.01, w) for w in weights]

    for _ in range(20):
        if remaining <= 1e-9 or not active:
            break
        total_w = sum(safe_weights[i] for i in active)
        consumed = 0.0
        saturated = []
        for i in list(active):
            share = remaining * safe_weights[i] / total_w
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
        candidates = [i for i in range(n) if widths[i] < maxs[i] - 1e-9]
        while remaining > 1e-8 and candidates:
            share = remaining / len(candidates)
            next_candidates = []
            for i in candidates:
                room = maxs[i] - widths[i]
                add = min(room, share)
                widths[i] += add
                remaining -= add
                if widths[i] < maxs[i] - 1e-9:
                    next_candidates.append(i)
            if len(next_candidates) == len(candidates) and share < 1e-10:
                break
            candidates = next_candidates

    total = sum(widths)
    return [w / total for w in widths] if total else [1 / n] * n


def column_stats(tbl: ET.Element, count: int) -> tuple[list[dict], list[str], list[str]]:
    hs = headers(tbl, count)
    roles = [role_for_header(h) for h in hs]
    vals: list[list[float]] = [[] for _ in range(count)]
    texts: list[list[str]] = [[] for _ in range(count)]

    rows = tbl.findall("w:tr", NS)
    for r_idx, row in enumerate(rows):
        for tc, start, end in row_cells_by_column(row, count):
            text = node_text(tc)
            if not text:
                continue
            pressure = visual_width(text)
            span = max(1, end - start)
            for col in range(start, end):
                # Headers are accounted for separately; body drives most pressure.
                if r_idx > 0:
                    vals[col].append(pressure / span)
                    texts[col].append(text)

    stats = []
    for i in range(count):
        body = vals[i]
        stats.append(
            {
                "header_pressure": visual_width(hs[i]),
                "avg_pressure": (sum(body) / len(body)) if body else 0.0,
                "p90_pressure": percentile(body, 0.90),
                "max_pressure": max(body) if body else 0.0,
                "body_samples": len(body),
                "max_text": max(texts[i], key=visual_width) if texts[i] else "",
            }
        )
    return stats, hs, roles


def recommended_fractions(stats: list[dict], roles: list[str], *, technical: bool = False) -> list[float]:
    n = len(stats)
    weights, mins, maxs = [], [], []
    for i, role in enumerate(roles):
        mn, mx, mult = _bounds(role, n)
        s = stats[i]
        pressure = (
            s["header_pressure"] * 0.70
            + s["avg_pressure"] * 0.85
            + s["p90_pressure"] * 0.85
            + s["max_pressure"] * 0.30
        )
        signal = 1.0 + math.log1p(max(0.0, pressure))
        weights.append(signal * mult)
        mins.append(mn)
        maxs.append(mx)

    # Technical deviation table: the third physical column may be a marker column.
    if technical and n == 7:
        mins[2] = min(mins[2], 0.035)
        maxs[2] = min(maxs[2], 0.07)
        weights[2] = min(weights[2], 0.35)
        roles[2] = "compact"

    return allocate_fractions(weights, mins, maxs)


def load_style_contract(path: Path | None) -> dict:
    if path is None or not path.is_file():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def expected_technical_grid(style_contract: dict) -> int | None:
    try:
        value = style_contract["table"]["technical_deviation"]["physical_grid_columns"]
        return int(value)
    except Exception:
        return None


def table_text_size_pt(style_contract: dict) -> float:
    for key in ("table_text_actual", "table_text_center_candidate"):
        try:
            v = style_contract["semantic_styles"][key]["size_pt"]
            return float(v)
        except Exception:
            pass
    return 12.0


@dataclass
class Issue:
    severity: str
    code: str
    table: int
    message: str
    column: int | None = None
    header: str | None = None
    evidence: dict | None = None
    suggestion: str | None = None
    requires_semantic_review: bool = False


def merge_candidates(tbl: ET.Element, table_index: int, hs: list[str], count: int) -> list[Issue]:
    rows = tbl.findall("w:tr", NS)
    if len(rows) < 3:
        return []

    # Map body rows to per-column text. Header = row 0.
    matrix: list[list[str]] = []
    for row in rows[1:]:
        vals = [""] * count
        for tc, start, end in row_cells_by_column(row, count):
            text = node_text(tc)
            for col in range(start, end):
                vals[col] = text
        matrix.append(vals)

    issues: list[Issue] = []
    for col in range(count):
        if not merge_friendly_header(hs[col]):
            continue
        start = 0
        while start < len(matrix):
            text = matrix[start][col]
            if not text:
                start += 1
                continue
            end = start + 1
            while end < len(matrix) and matrix[end][col] == text:
                end += 1
            run_len = end - start
            if run_len >= 2:
                # Require another column to vary across the group; otherwise it may be
                # duplicate rows rather than a grouping cell.
                varied = False
                for other in range(count):
                    if other == col:
                        continue
                    values = [matrix[r][other] for r in range(start, end)]
                    nonempty = [v for v in values if v]
                    if len(set(nonempty)) >= 2:
                        varied = True
                        break

                # Body row indexes in the full table: start+2 ... end+1.
                full_start = start + 1
                full_end = end
                already_merged = has_existing_vertical_merge(tbl, col, full_start, full_end)
                if varied and not already_merged:
                    issues.append(
                        Issue(
                            severity="review",
                            code="table-vertical-merge-candidate",
                            table=table_index,
                            column=col + 1,
                            header=hs[col],
                            message="同一分组列存在连续重复内容，可能应纵向合并；需要结合行语义确认。",
                            evidence={
                                "text": text[:120],
                                "body_row_start": start + 1,
                                "body_row_end": end,
                                "repeat_count": run_len,
                            },
                            suggestion="若这些行确属同一对象/类别/阶段，则纵向合并该列；若为独立记录则保持分开，并记录不合并理由。",
                            requires_semantic_review=True,
                        )
                    )
            start = end
    return issues


def longest_ascii_token(texts: Iterable[str]) -> str:
    best = ""
    for text in texts:
        for m in ASCII_TOKEN_RE.finditer(text or ""):
            token = m.group(0)
            if len(token) > len(best):
                best = token
    return best


def audit_table(
    tbl: ET.Element,
    table_index: int,
    printable_width: int,
    style_contract: dict,
) -> tuple[dict, list[Issue]]:
    issues: list[Issue] = []
    count = logical_column_count(tbl)
    if count <= 0:
        return {"table": table_index, "columns": 0}, issues

    widths = grid_widths(tbl, count)
    if not widths or len(widths) != count or any(w <= 0 for w in widths):
        issues.append(
            Issue(
                "error",
                "table-grid-width-invalid",
                table_index,
                "无法取得稳定的物理列宽；无法可靠判断列间协调性。",
                evidence={"grid_widths": widths},
                suggestion="写入明确的 tblGrid/tcW 列宽后重新检查。",
            )
        )
        actual = [1 / count] * count
        table_width = 0
    else:
        table_width = sum(widths)
        actual = [w / table_width for w in widths]

    technical = is_technical_deviation_table(tbl)
    stats, hs, roles = column_stats(tbl, count)
    rec = recommended_fractions(stats, roles[:], technical=technical)

    expected_grid = expected_technical_grid(style_contract)
    if technical and expected_grid is not None and count != expected_grid:
        issues.append(
            Issue(
                "error",
                "technical-deviation-physical-grid-mismatch",
                table_index,
                f"技术偏离表物理列数为 {count}，当前模板要求 {expected_grid} 个物理列。",
                evidence={"actual": count, "expected": expected_grid, "headers": hs},
                suggestion="按模板物理网格恢复列结构；语义表头可以跨列，但物理列数必须符合模板。",
            )
        )

    if table_width and printable_width and table_width > printable_width * 1.02:
        issues.append(
            Issue(
                "error",
                "table-width-overflow",
                table_index,
                "表格物理列宽总和超过当前节可打印宽度。",
                evidence={"table_width_twips": table_width, "printable_width_twips": printable_width},
                suggestion="重新分配列宽，不要靠无限拉伸页面或整体缩小字号解决。",
            )
        )

    # Table layout mode: autofit + long token can stretch Word unpredictably.
    layout_node = tbl.find("w:tblPr/w:tblLayout", NS)
    layout = layout_node.get(qn("type")) if layout_node is not None else "autofit"
    rows = tbl.findall("w:tr", NS)
    all_texts = [node_text(tc) for row in rows for tc in row.findall("w:tc", NS)]
    token = longest_ascii_token(all_texts)
    if token and layout != "fixed":
        # Approx 0.55em per ASCII char at table font size.
        size_pt = table_text_size_pt(style_contract)
        token_twips = len(token) * size_pt * 20 * 0.55
        if token_twips > max(printable_width * 0.72, (max(widths) if widths else printable_width / count) * 1.6):
            issues.append(
                Issue(
                    "error",
                    "table-autofit-page-stretch-risk",
                    table_index,
                    "表格使用自动调整列宽，且存在很长的不可断行字符串，可能把列或页面异常拉宽。",
                    evidence={"layout": layout, "token": token[:120], "estimated_token_twips": round(token_twips)},
                    suggestion="改为受控固定表格宽度，并对 URL/路径/长标识采用可断行或缩略展示策略。",
                )
            )

    if widths and table_width:
        size_pt = table_text_size_pt(style_contract)
        cjk_char_twips = max(180.0, size_pt * 20.0)
        median_pressure = statistics.median(
            [max(s["header_pressure"], s["avg_pressure"], s["p90_pressure"]) for s in stats]
        ) if stats else 0.0

        # Individual column allocation vs content pressure.
        for i in range(count):
            a = actual[i]
            r = rec[i] if i < len(rec) else 1 / count
            s = stats[i]
            effective = max(1.0, widths[i] - 180.0)
            char_capacity = effective / cjk_char_twips
            max_text = s["max_text"]
            has_cjk_body = bool(HAN_RE.search(max_text))

            if char_capacity < 1.65 and (has_cjk_body and visual_width(max_text) >= 4 or s["header_pressure"] >= 3):
                issues.append(
                    Issue(
                        "error",
                        "table-column-one-char-wrap-risk",
                        table_index,
                        "该列过窄，中文表头或正文存在接近“一字一行”的高风险。",
                        column=i + 1,
                        header=hs[i],
                        evidence={
                            "width_twips": widths[i],
                            "estimated_cjk_chars_per_line": round(char_capacity, 2),
                            "max_text": max_text[:120],
                        },
                        suggestion="增加该列宽度；优先从内容压力低、明显过宽的列回收空间。",
                    )
                )
            elif char_capacity < 2.25 and s["header_pressure"] >= 4:
                issues.append(
                    Issue(
                        "warning",
                        "table-header-cramped",
                        table_index,
                        "表头相对列宽偏挤，渲染后可能出现突兀的多行拆字。",
                        column=i + 1,
                        header=hs[i],
                        evidence={
                            "width_twips": widths[i],
                            "estimated_cjk_chars_per_line": round(char_capacity, 2),
                            "header_pressure": round(s["header_pressure"], 2),
                        },
                        suggestion="结合表头字数和正文内容重新平衡列宽。",
                    )
                )

            pressure = max(s["header_pressure"], s["avg_pressure"], s["p90_pressure"])
            if r > 0 and a < r * 0.46 and pressure >= max(3.0, median_pressure * 0.85):
                issues.append(
                    Issue(
                        "error",
                        "table-column-severely-underallocated",
                        table_index,
                        "当前列宽明显小于其表头/正文内容压力所需要的合理宽度。",
                        column=i + 1,
                        header=hs[i],
                        evidence={
                            "actual_fraction": round(a, 4),
                            "recommended_fraction": round(r, 4),
                            "avg_pressure": round(s["avg_pressure"], 2),
                            "p90_pressure": round(s["p90_pressure"], 2),
                            "max_pressure": round(s["max_pressure"], 2),
                        },
                        suggestion="按表头与正文内容压力重新分配宽度，不要让长文本列长期拥挤。",
                    )
                )
            elif r > 0 and a < r * 0.66 and pressure >= max(3.0, median_pressure):
                issues.append(
                    Issue(
                        "warning",
                        "table-column-underallocated",
                        table_index,
                        "当前列偏窄，与该列内容量不匹配，容易造成视觉拥挤。",
                        column=i + 1,
                        header=hs[i],
                        evidence={"actual_fraction": round(a, 4), "recommended_fraction": round(r, 4)},
                        suggestion="增加该列宽度并从低压力列回收空间。",
                    )
                )

        under = [
            i for i in range(count)
            if rec[i] > 0 and actual[i] < rec[i] * 0.72
        ]
        for i in range(count):
            r = rec[i]
            if r <= 0:
                continue
            s = stats[i]
            pressure = max(s["header_pressure"], s["avg_pressure"], s["p90_pressure"])
            if actual[i] > r * 1.75 and under and pressure <= max(6.0, median_pressure * 0.90):
                issues.append(
                    Issue(
                        "warning",
                        "table-column-overallocated",
                        table_index,
                        "该列明显偏宽，而其他列存在宽度不足，整体视觉比例可能突兀。",
                        column=i + 1,
                        header=hs[i],
                        evidence={
                            "actual_fraction": round(actual[i], 4),
                            "recommended_fraction": round(r, 4),
                            "underallocated_columns": [x + 1 for x in under],
                        },
                        suggestion="压缩低内容压力列，把空间让给正文更长的列。",
                    )
                )

        # Adjacent abrupt imbalance not justified by pressure. Skip marker column of
        # the approved 7-grid technical deviation table.
        for i in range(count - 1):
            if technical and count == 7 and (i == 1 or i == 2):
                continue
            w1, w2 = actual[i], actual[i + 1]
            ratio = max(w1, w2) / max(min(w1, w2), 1e-9)
            p1 = max(stats[i]["header_pressure"], stats[i]["avg_pressure"], stats[i]["p90_pressure"], 1.0)
            p2 = max(stats[i + 1]["header_pressure"], stats[i + 1]["avg_pressure"], stats[i + 1]["p90_pressure"], 1.0)
            pressure_ratio = max(p1, p2) / max(min(p1, p2), 1e-9)
            if ratio >= 3.2 and pressure_ratio <= 1.55:
                issues.append(
                    Issue(
                        "warning",
                        "table-adjacent-width-abrupt",
                        table_index,
                        "相邻两列宽度反差很大，但内容压力差异不足以解释这种比例，视觉上可能明显不协调。",
                        column=i + 1,
                        header=f"{hs[i]} | {hs[i + 1]}",
                        evidence={
                            "left_fraction": round(w1, 4),
                            "right_fraction": round(w2, 4),
                            "width_ratio": round(ratio, 2),
                            "content_pressure_ratio": round(pressure_ratio, 2),
                        },
                        suggestion="重新平衡相邻列宽，让列宽变化与表头和正文内容量相匹配。",
                    )
                )

    issues.extend(merge_candidates(tbl, table_index, hs, count))

    report = {
        "table": table_index,
        "technical_deviation": technical,
        "columns": count,
        "headers": hs,
        "roles": roles,
        "printable_width_twips": printable_width,
        "table_width_twips": table_width,
        "layout": layout,
        "column_widths_twips": widths,
        "actual_fractions": [round(x, 4) for x in actual],
        "recommended_fractions": [round(x, 4) for x in rec],
        "column_stats": stats,
    }
    return report, issues


def audit(docx: Path, style_json: Path | None = None) -> dict:
    style_contract = load_style_contract(style_json)
    with zipfile.ZipFile(docx, "r") as z:
        root = ET.fromstring(z.read("word/document.xml"))
    body = root.find("w:body", NS)
    if body is None:
        raise ValueError("DOCX document.xml lacks w:body")

    section_widths = table_section_widths(body)
    tables = root.findall(".//w:tbl", NS)
    table_reports = []
    issues: list[Issue] = []

    for idx, tbl in enumerate(tables, 1):
        report, found = audit_table(
            tbl,
            idx,
            section_widths.get(id(tbl), 8306),
            style_contract,
        )
        table_reports.append(report)
        issues.extend(found)

    counts = {"error": 0, "warning": 0, "review": 0, "info": 0}
    for issue in issues:
        counts[issue.severity] = counts.get(issue.severity, 0) + 1

    return {
        "input": str(docx),
        "template_style_json": str(style_json) if style_json else None,
        "status": "failed" if counts["error"] else ("review_required" if counts["review"] else "passed"),
        "summary": {
            "tables": len(tables),
            "issues": len(issues),
            "severity": counts,
        },
        "tables": table_reports,
        "issues": [asdict(i) for i in issues],
        "visual_review_required": True,
        "note": (
            "列宽协调性最终仍需检查最新 DOCX 的真实渲染页面。"
            "review 级纵向合并候选必须由 Agent 结合语义决定是否合并，不能机械自动合并。"
        ),
    }


def main() -> int:
    ap = argparse.ArgumentParser(description="Audit DOCX table width balance, wrapping risk and merge candidates.")
    ap.add_argument("docx", type=Path)
    ap.add_argument("--template-style-json", type=Path)
    ap.add_argument("--json-out", type=Path)
    args = ap.parse_args()

    if not args.docx.is_file():
        ap.error("DOCX not found")
    try:
        result = audit(args.docx, args.template_style_json)
    except (zipfile.BadZipFile, KeyError, ET.ParseError, ValueError) as exc:
        payload = {"status": "failed", "error": str(exc)}
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 3

    payload = json.dumps(result, ensure_ascii=False, indent=2)
    if args.json_out:
        args.json_out.parent.mkdir(parents=True, exist_ok=True)
        args.json_out.write_text(payload + "\n", encoding="utf-8")
    print(payload)
    return 2 if result["summary"]["severity"]["error"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
