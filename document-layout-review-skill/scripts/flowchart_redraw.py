#!/usr/bin/env python3
"""Redraw a structured flowchart spec as a clean SVG.

The script intentionally redraws from semantic nodes/edges rather than preserving
coordinates from a source bitmap. It is a dependency-free portable fallback.
Complex graphs should use a stronger layout engine such as ELK when available.
"""
from __future__ import annotations

import argparse
from copy import deepcopy
import json
import math
import re
from punctuation_context import normalize_fragment as normalize_punctuation_fragment
from punctuation_context import normalize_fragment
import shlex
import shutil
import subprocess
import sys
import tempfile
from collections import defaultdict, deque
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable
from xml.sax.saxutils import escape

# Local import: flowchart_audit.py lives next to this script.
try:
    from flowchart_audit import audit_one
except Exception:  # pragma: no cover - usable as imported module too
    audit_one = None  # type: ignore[assignment]


FONT_SIZE = 14
LINE_HEIGHT = 20
PAD_X = 18
PAD_Y = 14
MIN_NODE_W = 120
MAX_NODE_W = 260
MIN_NODE_H = 54
MARGIN = 36
CROSS_GAP = 56
MAIN_GAP = 88
GROUP_PAD = 22


@dataclass(frozen=True)
class NodeBox:
    node_id: str
    node_type: str
    text: str
    lines: tuple[str, ...]
    x: float
    y: float
    width: float
    height: float
    group: str | None = None

    @property
    def cx(self) -> float:
        return self.x + self.width / 2

    @property
    def cy(self) -> float:
        return self.y + self.height / 2


@dataclass(frozen=True)
class EdgePath:
    source: str
    target: str
    label: str
    points: tuple[tuple[float, float], ...]
    label_point: tuple[float, float]


def visual_units(text: str) -> float:
    units = 0.0
    for ch in text:
        if ch == "\t":
            units += 4
        elif ord(ch) < 128:
            units += 0.58
        else:
            units += 1.0
    return units


def wrap_text(text: str, max_units: float = 16.0) -> tuple[str, ...]:
    text = re.sub(r"[ \t\u3000]+", " ", str(text or "").strip())
    if not text:
        return ("",)
    paragraphs = text.splitlines() or [text]
    out: list[str] = []
    for paragraph in paragraphs:
        if visual_units(paragraph) <= max_units:
            out.append(paragraph)
            continue
        current = ""
        current_units = 0.0
        for ch in paragraph:
            u = visual_units(ch)
            # Prefer breaking before ASCII words only when a space is available.
            if current and current_units + u > max_units:
                split_at = current.rfind(" ")
                if split_at >= max(2, len(current) // 3):
                    head, tail = current[:split_at].rstrip(), current[split_at + 1 :].lstrip()
                    out.append(head)
                    current = tail + ch
                    current_units = visual_units(current)
                else:
                    out.append(current.rstrip())
                    current = ch
                    current_units = u
            else:
                current += ch
                current_units += u
        if current or not out:
            out.append(current.rstrip())
    return tuple(line if line else " " for line in out)


def node_dimensions(node: dict[str, Any]) -> tuple[tuple[str, ...], float, float]:
    lines = wrap_text(str(node.get("text") or ""))
    max_units = max((visual_units(line) for line in lines), default=4)
    width = max(MIN_NODE_W, min(MAX_NODE_W, max_units * FONT_SIZE + PAD_X * 2))
    if node.get("type") == "decision":
        width = max(width, 170)
    height = max(MIN_NODE_H, len(lines) * LINE_HEIGHT + PAD_Y * 2)
    if node.get("type") == "decision":
        height = max(height, 80)
    return lines, width, height


def tarjan_scc(node_ids: list[str], outgoing: dict[str, list[str]]) -> list[list[str]]:
    index = 0
    stack: list[str] = []
    on_stack: set[str] = set()
    indices: dict[str, int] = {}
    low: dict[str, int] = {}
    components: list[list[str]] = []

    def visit(v: str) -> None:
        nonlocal index
        indices[v] = index
        low[v] = index
        index += 1
        stack.append(v)
        on_stack.add(v)
        for w in outgoing.get(v, []):
            if w not in indices:
                visit(w)
                low[v] = min(low[v], low[w])
            elif w in on_stack:
                low[v] = min(low[v], indices[w])
        if low[v] == indices[v]:
            comp: list[str] = []
            while True:
                w = stack.pop()
                on_stack.remove(w)
                comp.append(w)
                if w == v:
                    break
            components.append(comp)

    for node_id in node_ids:
        if node_id not in indices:
            visit(node_id)
    return components


def assign_ranks(node_ids: list[str], edges: list[dict[str, Any]]) -> tuple[dict[str, int], dict[str, int]]:
    outgoing: dict[str, list[str]] = defaultdict(list)
    for edge in edges:
        src, dst = str(edge.get("from")), str(edge.get("to"))
        if src in node_ids and dst in node_ids:
            outgoing[src].append(dst)

    components = tarjan_scc(node_ids, outgoing)
    component_of: dict[str, int] = {}
    for idx, comp in enumerate(components):
        for nid in comp:
            component_of[nid] = idx

    comp_out: dict[int, set[int]] = defaultdict(set)
    indegree: dict[int, int] = {i: 0 for i in range(len(components))}
    for src in node_ids:
        c_src = component_of[src]
        for dst in outgoing.get(src, []):
            c_dst = component_of[dst]
            if c_src != c_dst and c_dst not in comp_out[c_src]:
                comp_out[c_src].add(c_dst)
                indegree[c_dst] += 1

    queue = deque(sorted((cid for cid, deg in indegree.items() if deg == 0)))
    comp_rank = {cid: 0 for cid in indegree}
    while queue:
        cid = queue.popleft()
        for nxt in sorted(comp_out.get(cid, set())):
            comp_rank[nxt] = max(comp_rank[nxt], comp_rank[cid] + 1)
            indegree[nxt] -= 1
            if indegree[nxt] == 0:
                queue.append(nxt)

    ranks = {nid: comp_rank[component_of[nid]] for nid in node_ids}
    # Stable index inside SCC helps route cycle edges on outer corridors.
    scc_index: dict[str, int] = {}
    for comp in components:
        ordered = sorted(comp, key=node_ids.index)
        for i, nid in enumerate(ordered):
            scc_index[nid] = i
    return ranks, scc_index


def layout_nodes(spec: dict[str, Any]) -> tuple[dict[str, NodeBox], float, float, dict[str, int]]:
    nodes = [n for n in spec.get("nodes", []) if isinstance(n, dict) and n.get("id")]
    edges = [e for e in spec.get("edges", []) if isinstance(e, dict)]
    node_ids = [str(n["id"]) for n in nodes]
    direction = spec.get("direction") if spec.get("direction") in ("TB", "LR") else "TB"
    ranks, scc_index = assign_ranks(node_ids, edges)
    levels: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for node in nodes:
        levels[ranks[str(node["id"])]].append(node)

    # Preserve semantic source order inside each level, then lightly cluster by group.
    source_pos = {nid: i for i, nid in enumerate(node_ids)}
    for level in levels.values():
        level.sort(key=lambda n: (str(n.get("group") or ""), source_pos[str(n["id"])]))

    dimensions = {str(n["id"]): node_dimensions(n) for n in nodes}
    ordered_levels = sorted(levels)
    main_sizes: dict[int, float] = {}
    cross_sizes: dict[int, float] = {}
    for rank in ordered_levels:
        items = levels[rank]
        if direction == "TB":
            main_sizes[rank] = max(dimensions[str(n["id"])][2] for n in items)
            cross_sizes[rank] = sum(dimensions[str(n["id"])][1] for n in items) + CROSS_GAP * max(0, len(items) - 1)
        else:
            main_sizes[rank] = max(dimensions[str(n["id"])][1] for n in items)
            cross_sizes[rank] = sum(dimensions[str(n["id"])][2] for n in items) + CROSS_GAP * max(0, len(items) - 1)

    max_cross = max(cross_sizes.values(), default=320)
    main_cursor = MARGIN
    main_coord: dict[int, float] = {}
    for rank in ordered_levels:
        main_coord[rank] = main_cursor
        main_cursor += main_sizes[rank] + MAIN_GAP

    boxes: dict[str, NodeBox] = {}
    for rank in ordered_levels:
        items = levels[rank]
        cross_cursor = MARGIN + (max_cross - cross_sizes[rank]) / 2
        for node in items:
            nid = str(node["id"])
            lines, width, height = dimensions[nid]
            if direction == "TB":
                x, y = cross_cursor, main_coord[rank]
                cross_cursor += width + CROSS_GAP
            else:
                x, y = main_coord[rank], cross_cursor
                cross_cursor += height + CROSS_GAP
            boxes[nid] = NodeBox(
                node_id=nid,
                node_type=str(node.get("type") or "process"),
                text=str(node.get("text") or ""),
                lines=lines,
                x=x,
                y=y,
                width=width,
                height=height,
                group=(str(node.get("group")) if node.get("group") else None),
            )

    if direction == "TB":
        width = max_cross + MARGIN * 2
        height = main_cursor - MAIN_GAP + MARGIN
    else:
        width = main_cursor - MAIN_GAP + MARGIN
        height = max_cross + MARGIN * 2
    return boxes, max(360, width), max(220, height), scc_index


def port(box: NodeBox, side: str) -> tuple[float, float]:
    if side == "top":
        return box.cx, box.y
    if side == "bottom":
        return box.cx, box.y + box.height
    if side == "left":
        return box.x, box.cy
    return box.x + box.width, box.cy


def compress_points(points: Iterable[tuple[float, float]]) -> tuple[tuple[float, float], ...]:
    out: list[tuple[float, float]] = []
    for p in points:
        if out and math.isclose(out[-1][0], p[0]) and math.isclose(out[-1][1], p[1]):
            continue
        out.append(p)
    # Drop collinear middle points.
    changed = True
    while changed and len(out) >= 3:
        changed = False
        nxt = [out[0]]
        for i in range(1, len(out) - 1):
            a, b, c = nxt[-1], out[i], out[i + 1]
            if (math.isclose(a[0], b[0]) and math.isclose(b[0], c[0])) or (math.isclose(a[1], b[1]) and math.isclose(b[1], c[1])):
                changed = True
                continue
            nxt.append(b)
        nxt.append(out[-1])
        out = nxt
    return tuple(out)


def midpoint_of_longest_segment(points: tuple[tuple[float, float], ...]) -> tuple[float, float]:
    if len(points) < 2:
        return points[0] if points else (0, 0)
    best = (0.0, points[0], points[1])
    for a, b in zip(points, points[1:]):
        d = abs(a[0] - b[0]) + abs(a[1] - b[1])
        if d > best[0]:
            best = (d, a, b)
    _, a, b = best
    return ((a[0] + b[0]) / 2, (a[1] + b[1]) / 2)


def route_edges(spec: dict[str, Any], boxes: dict[str, NodeBox], canvas_w: float, canvas_h: float) -> tuple[list[EdgePath], float, float]:
    direction = spec.get("direction") if spec.get("direction") in ("TB", "LR") else "TB"
    edges = [e for e in spec.get("edges", []) if isinstance(e, dict)]
    paths: list[EdgePath] = []
    outer_count = 0
    extra_w = 0.0
    extra_h = 0.0

    for edge in edges:
        src_id, dst_id = str(edge.get("from")), str(edge.get("to"))
        if src_id not in boxes or dst_id not in boxes:
            continue
        src, dst = boxes[src_id], boxes[dst_id]
        label = str(edge.get("label") or "")
        if direction == "TB":
            clearly_forward = dst.y >= src.y + src.height + 8
            clearly_backward = src.y >= dst.y + dst.height + 8
            if clearly_forward:
                start = port(src, "bottom")
                end = port(dst, "top")
                mid_y = (start[1] + end[1]) / 2
                points = compress_points((start, (start[0], mid_y), (end[0], mid_y), end))
            elif clearly_backward:
                outer_count += 1
                corridor = canvas_w + 32 + outer_count * 24
                extra_w = max(extra_w, corridor - canvas_w + MARGIN)
                start = port(src, "right")
                end = port(dst, "right")
                points = compress_points((start, (corridor, start[1]), (corridor, end[1]), end))
            else:
                # Same-rank branch: route below both nodes so the arrow never crosses either box.
                outer_count += 1
                corridor_y = max(src.y + src.height, dst.y + dst.height) + 30 + outer_count * 18
                extra_h = max(extra_h, corridor_y - canvas_h + MARGIN)
                start = port(src, "bottom")
                end = port(dst, "bottom")
                points = compress_points((start, (start[0], corridor_y), (end[0], corridor_y), end))
        else:
            clearly_forward = dst.x >= src.x + src.width + 8
            clearly_backward = src.x >= dst.x + dst.width + 8
            if clearly_forward:
                start = port(src, "right")
                end = port(dst, "left")
                mid_x = (start[0] + end[0]) / 2
                points = compress_points((start, (mid_x, start[1]), (mid_x, end[1]), end))
            elif clearly_backward:
                outer_count += 1
                corridor = canvas_h + 32 + outer_count * 24
                extra_h = max(extra_h, corridor - canvas_h + MARGIN)
                start = port(src, "bottom")
                end = port(dst, "bottom")
                points = compress_points((start, (start[0], corridor), (end[0], corridor), end))
            else:
                # Same-column branch: route to the right of both nodes.
                outer_count += 1
                corridor_x = max(src.x + src.width, dst.x + dst.width) + 30 + outer_count * 18
                extra_w = max(extra_w, corridor_x - canvas_w + MARGIN)
                start = port(src, "right")
                end = port(dst, "right")
                points = compress_points((start, (corridor_x, start[1]), (corridor_x, end[1]), end))
        lp = midpoint_of_longest_segment(points)
        paths.append(EdgePath(src_id, dst_id, label, points, lp))
    return paths, canvas_w + extra_w, canvas_h + extra_h


def svg_node(box: NodeBox) -> str:
    x, y, w, h = box.x, box.y, box.width, box.height
    node_type = box.node_type
    stroke = "#334155"
    fill = "#ffffff"
    if node_type in ("start", "end"):
        stroke, fill = "#047857", "#ecfdf5"
    elif node_type == "decision":
        stroke, fill = "#c2410c", "#fff7ed"
    elif node_type == "subprocess":
        stroke, fill = "#1d4ed8", "#eff6ff"
    elif node_type == "document":
        stroke, fill = "#6d28d9", "#f5f3ff"

    if node_type == "decision":
        shape = f'<polygon points="{x+w/2:.1f},{y:.1f} {x+w:.1f},{y+h/2:.1f} {x+w/2:.1f},{y+h:.1f} {x:.1f},{y+h/2:.1f}" fill="{fill}" stroke="{stroke}" stroke-width="2"/>'
    else:
        radius = h / 2 if node_type in ("start", "end") else 8
        shape = f'<rect x="{x:.1f}" y="{y:.1f}" width="{w:.1f}" height="{h:.1f}" rx="{radius:.1f}" fill="{fill}" stroke="{stroke}" stroke-width="2"/>'
        if node_type == "subprocess":
            shape += f'<line x1="{x+10:.1f}" y1="{y:.1f}" x2="{x+10:.1f}" y2="{y+h:.1f}" stroke="{stroke}" stroke-width="1.5"/><line x1="{x+w-10:.1f}" y1="{y:.1f}" x2="{x+w-10:.1f}" y2="{y+h:.1f}" stroke="{stroke}" stroke-width="1.5"/>'

    start_y = box.cy - (len(box.lines) - 1) * LINE_HEIGHT / 2
    tspans = []
    for i, line in enumerate(box.lines):
        dy = 0 if i == 0 else LINE_HEIGHT
        tspans.append(f'<tspan x="{box.cx:.1f}" dy="{dy}">{escape(line)}</tspan>')
    text = f'<text x="{box.cx:.1f}" y="{start_y:.1f}" text-anchor="middle" dominant-baseline="middle" font-family="Microsoft YaHei, PingFang SC, Noto Sans CJK SC, Arial, sans-serif" font-size="{FONT_SIZE}" fill="#0f172a">{"".join(tspans)}</text>'
    return shape + text


def group_shapes(spec: dict[str, Any], boxes: dict[str, NodeBox]) -> str:
    labels: dict[str, str] = {}
    for group in spec.get("groups", []) if isinstance(spec.get("groups"), list) else []:
        if isinstance(group, dict) and group.get("id"):
            labels[str(group["id"])] = str(group.get("label") or group["id"])
    grouped: dict[str, list[NodeBox]] = defaultdict(list)
    for box in boxes.values():
        if box.group:
            grouped[box.group].append(box)
    parts: list[str] = []
    for gid, items in grouped.items():
        min_x = min(i.x for i in items) - GROUP_PAD
        min_y = min(i.y for i in items) - GROUP_PAD - 12
        max_x = max(i.x + i.width for i in items) + GROUP_PAD
        max_y = max(i.y + i.height for i in items) + GROUP_PAD
        parts.append(f'<rect x="{min_x:.1f}" y="{min_y:.1f}" width="{max_x-min_x:.1f}" height="{max_y-min_y:.1f}" rx="10" fill="none" stroke="#94a3b8" stroke-width="1.5" stroke-dasharray="6 5"/>')
        parts.append(f'<text x="{min_x+10:.1f}" y="{min_y+17:.1f}" font-family="Microsoft YaHei, PingFang SC, Noto Sans CJK SC, Arial, sans-serif" font-size="12" fill="#475569">{escape(labels.get(gid, gid))}</text>')
    return "".join(parts)


def render_svg_fallback(spec: dict[str, Any]) -> str:
    boxes, width, height, _ = layout_nodes(spec)
    paths, width, height = route_edges(spec, boxes, width, height)
    title = str(spec.get("title") or "流程图")

    edge_parts: list[str] = []
    for idx, path in enumerate(paths):
        points = " ".join(f"{x:.1f},{y:.1f}" for x, y in path.points)
        edge_parts.append(f'<polyline id="edge-{idx+1}" points="{points}" fill="none" stroke="#64748b" stroke-width="2" stroke-linejoin="round" stroke-linecap="round" marker-end="url(#arrow)"/>')
        if path.label.strip():
            lx, ly = path.label_point
            label = escape(path.label.strip())
            pad = max(16, visual_units(path.label.strip()) * 6.5 + 10)
            edge_parts.append(f'<rect x="{lx-pad/2:.1f}" y="{ly-10:.1f}" width="{pad:.1f}" height="20" rx="4" fill="#ffffff" fill-opacity="0.92"/>')
            edge_parts.append(f'<text x="{lx:.1f}" y="{ly+1:.1f}" text-anchor="middle" dominant-baseline="middle" font-family="Microsoft YaHei, PingFang SC, Noto Sans CJK SC, Arial, sans-serif" font-size="12" fill="#475569">{label}</text>')

    nodes = "".join(svg_node(box) for box in boxes.values())
    groups = group_shapes(spec, boxes)
    defs = '<defs><marker id="arrow" markerWidth="10" markerHeight="10" refX="8.5" refY="5" orient="auto" markerUnits="strokeWidth"><path d="M0,0 L10,5 L0,10 z" fill="#64748b"/></marker></defs>'
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" role="img" aria-label="{escape(title)}" '
        f'width="{math.ceil(width)}" height="{math.ceil(height)}" viewBox="0 0 {math.ceil(width)} {math.ceil(height)}">'
        f'{defs}<rect width="100%" height="100%" fill="white"/>{groups}{"".join(edge_parts)}{nodes}</svg>\n'
    )



def normalize_contextual_punctuation(value: str) -> str:
    """Use full-width punctuation for Chinese and half-width for English/code/formulas."""
    normalized, _kind = normalize_fragment(str(value or ""))
    return normalized


def dot_escape(value: str) -> str:
    return value.replace("\\", r"\\").replace('"', r'\"')


def graph_is_complex(spec: dict[str, Any]) -> bool:
    nodes = [n for n in spec.get("nodes", []) if isinstance(n, dict)]
    edges = [e for e in spec.get("edges", []) if isinstance(e, dict)]
    decisions = sum(1 for n in nodes if n.get("type") == "decision")
    node_ids = [str(n.get("id")) for n in nodes if n.get("id")]
    source_pos = {nid: i for i, nid in enumerate(node_ids)}
    has_back_edge = any(
        str(e.get("from")) in source_pos and str(e.get("to")) in source_pos
        and source_pos[str(e.get("to"))] <= source_pos[str(e.get("from"))]
        for e in edges
    )
    out_degree: dict[str, int] = defaultdict(int)
    for e in edges:
        out_degree[str(e.get("from"))] += 1
    return decisions >= 2 or has_back_edge or any(v > 2 for v in out_degree.values()) or len(edges) > len(nodes) + 1


def graphviz_dot(spec: dict[str, Any]) -> str:
    direction = spec.get("direction") if spec.get("direction") in ("TB", "LR") else "TB"
    nodes = [n for n in spec.get("nodes", []) if isinstance(n, dict) and n.get("id")]
    edges = [e for e in spec.get("edges", []) if isinstance(e, dict)]
    node_ids = [str(n["id"]) for n in nodes]
    source_pos = {nid: i for i, nid in enumerate(node_ids)}
    lines = [
        "digraph G {",
        f'  graph [rankdir={direction}, splines=ortho, nodesep=0.55, ranksep=0.65, pad=0.22, margin=0.04, bgcolor="white", outputorder=edgesfirst];',
        '  node [shape=box, style="rounded,filled", fillcolor="white", color="#475569", fontname="Microsoft YaHei, Noto Sans CJK SC, sans-serif", fontsize=14, margin="0.18,0.12", penwidth=1.5];',
        '  edge [color="#64748b", penwidth=1.6, arrowsize=0.8, fontname="Microsoft YaHei, Noto Sans CJK SC, sans-serif", fontsize=11];',
    ]
    for node in nodes:
        nid = str(node["id"])
        node_type = str(node.get("type") or "process")
        text = normalize_contextual_punctuation(str(node.get("text") or ""))
        wrapped = wrap_text(text, max_units=18.0)
        label = r"\n".join(dot_escape(line) for line in wrapped)
        attrs: list[str] = [f'label="{label}"']
        if node_type in ("start", "end"):
            attrs.extend(['shape=oval', 'fillcolor="#ecfdf5"', 'color="#047857"'])
        elif node_type == "decision":
            attrs.extend(['shape=diamond', 'fillcolor="#fff7ed"', 'color="#c2410c"', 'margin="0.12,0.08"'])
        elif node_type == "subprocess":
            attrs.extend(['shape=box', 'peripheries=2', 'fillcolor="#eff6ff"', 'color="#1d4ed8"'])
        elif node_type == "document":
            attrs.extend(['shape=note', 'fillcolor="#f5f3ff"', 'color="#6d28d9"'])
        lines.append(f'  "{dot_escape(nid)}" [{", ".join(attrs)}];')
    negative_words = ("否", "整改", "回归", "重新", "补证", "返工", "重试", "退回")
    for edge in edges:
        src, dst = str(edge.get("from")), str(edge.get("to"))
        if src not in source_pos or dst not in source_pos:
            continue
        label = normalize_contextual_punctuation(str(edge.get("label") or "").strip())
        attrs: list[str] = []
        if label:
            attrs.append(f'xlabel="{dot_escape(label)}"')
        # Back/loop edges must not distort the main rank order; let dot route them
        # around the outer corridor instead.
        is_back_edge = source_pos.get(dst, 0) <= source_pos.get(src, 0)
        if is_back_edge:
            attrs.append("constraint=false")
        if any(word in label for word in negative_words) or is_back_edge:
            attrs.extend(['color="#c2410c"', 'fontcolor="#9a3412"'])
        lines.append(f'  "{dot_escape(src)}" -> "{dot_escape(dst)}"' + (f' [{", ".join(attrs)}]' if attrs else '') + ';')
    lines.append("}")
    return "\n".join(lines) + "\n"


def parse_graphviz_plain(plain: str) -> tuple[dict[str, tuple[float, float, float, float]], list[tuple[str, str, list[tuple[float, float]]]]]:
    nodes: dict[str, tuple[float, float, float, float]] = {}
    edges: list[tuple[str, str, list[tuple[float, float]]]] = []
    for raw in plain.splitlines():
        if not raw.strip():
            continue
        parts = shlex.split(raw)
        if not parts:
            continue
        if parts[0] == "node" and len(parts) >= 6:
            nodes[parts[1]] = tuple(map(float, parts[2:6]))  # cx, cy, width, height
        elif parts[0] == "edge" and len(parts) >= 4:
            n = int(parts[3])
            coords = [float(x) for x in parts[4:4 + 2*n]]
            pts = [(coords[i], coords[i+1]) for i in range(0, len(coords), 2)]
            edges.append((parts[1], parts[2], pts))
    return nodes, edges


def segment_hits_rect(a: tuple[float, float], b: tuple[float, float], rect: tuple[float, float, float, float], eps: float = 0.025) -> bool:
    cx, cy, w, h = rect
    left, right = cx - w/2 + eps, cx + w/2 - eps
    bottom, top = cy - h/2 + eps, cy + h/2 - eps
    if math.isclose(a[0], b[0], abs_tol=1e-7):
        x = a[0]
        lo, hi = sorted((a[1], b[1]))
        return left < x < right and max(lo, bottom) < min(hi, top)
    if math.isclose(a[1], b[1], abs_tol=1e-7):
        y = a[1]
        lo, hi = sorted((a[0], b[0]))
        return bottom < y < top and max(lo, left) < min(hi, right)
    # Orthogonal routing is required. A diagonal segment is itself a layout defect.
    return False


def segments_cross(a1: tuple[float, float], a2: tuple[float, float], b1: tuple[float, float], b2: tuple[float, float], eps: float = 1e-6) -> bool:
    a_vert = math.isclose(a1[0], a2[0], abs_tol=eps)
    b_vert = math.isclose(b1[0], b2[0], abs_tol=eps)
    if a_vert == b_vert:
        return False
    if a_vert:
        vx, hy = a1[0], b1[1]
        ay1, ay2 = sorted((a1[1], a2[1]))
        bx1, bx2 = sorted((b1[0], b2[0]))
        return ay1 + eps < hy < ay2 - eps and bx1 + eps < vx < bx2 - eps
    return segments_cross(b1, b2, a1, a2, eps)


def validate_graphviz_geometry(plain: str) -> dict[str, Any]:
    nodes, edges = parse_graphviz_plain(plain)
    issues: list[dict[str, Any]] = []
    for src, dst, pts in edges:
        for i, (a, b) in enumerate(zip(pts, pts[1:]), 1):
            if not (math.isclose(a[0], b[0], abs_tol=1e-7) or math.isclose(a[1], b[1], abs_tol=1e-7)):
                issues.append({"code": "diagonal-connector", "severity": "error", "edge": f"{src}->{dst}", "segment": i})
            for nid, rect in nodes.items():
                if nid in (src, dst):
                    continue
                if segment_hits_rect(a, b, rect):
                    issues.append({"code": "connector-through-node", "severity": "error", "edge": f"{src}->{dst}", "node": nid, "segment": i})
    for i, (s1, t1, p1) in enumerate(edges):
        for s2, t2, p2 in edges[i+1:]:
            if {s1, t1} & {s2, t2}:
                continue
            found = False
            for a1, a2 in zip(p1, p1[1:]):
                for b1, b2 in zip(p2, p2[1:]):
                    if segments_cross(a1, a2, b1, b2):
                        issues.append({"code": "connector-crossing", "severity": "error", "edges": [f"{s1}->{t1}", f"{s2}->{t2}"]})
                        found = True
                        break
                if found:
                    break
    return {
        "engine": "graphviz-dot",
        "node_count": len(nodes),
        "edge_count": len(edges),
        "issues": issues,
        "error_count": sum(1 for i in issues if i.get("severity") == "error"),
    }


def render_with_graphviz(spec: dict[str, Any]) -> tuple[str, dict[str, Any], str]:
    dot_bin = shutil.which("dot")
    if not dot_bin:
        raise RuntimeError("GRAPHVIZ_DOT_UNAVAILABLE")
    source = graphviz_dot(spec)
    with tempfile.TemporaryDirectory(prefix="flowchart-redraw-") as tmp:
        dot_path = Path(tmp) / "flow.dot"
        dot_path.write_text(source, encoding="utf-8")
        svg_run = subprocess.run([dot_bin, "-Tsvg", str(dot_path)], capture_output=True, text=True, encoding="utf-8")
        if svg_run.returncode != 0:
            raise RuntimeError("GRAPHVIZ_SVG_FAILED:" + svg_run.stderr.strip())
        plain_run = subprocess.run([dot_bin, "-Tplain", str(dot_path)], capture_output=True, text=True, encoding="utf-8")
        if plain_run.returncode != 0:
            raise RuntimeError("GRAPHVIZ_PLAIN_FAILED:" + plain_run.stderr.strip())
        geometry = validate_graphviz_geometry(plain_run.stdout)
        return svg_run.stdout, geometry, source


def render_svg(spec: dict[str, Any], engine: str = "auto", allow_crossings: bool = False) -> tuple[str, dict[str, Any], str | None]:
    complex_graph = graph_is_complex(spec)
    chosen = engine
    if engine == "auto":
        chosen = "graphviz" if shutil.which("dot") else "fallback"
    if chosen == "graphviz":
        svg, geometry, dot_source = render_with_graphviz(spec)
        if geometry["error_count"] and not allow_crossings:
            raise RuntimeError("FLOWCHART_GEOMETRY_REVIEW_FAILED:" + json.dumps(geometry, ensure_ascii=False))
        return svg, geometry, dot_source
    if complex_graph:
        raise RuntimeError("FLOWCHART_COMPLEX_LAYOUT_ENGINE_REQUIRED: install Graphviz dot or use host ELK/elkjs")
    svg = render_svg_fallback(spec)
    return svg, {"engine": "fallback", "issues": [], "error_count": 0}, None

def normalize_spec_punctuation(spec: dict[str, Any]) -> tuple[dict[str, Any], list[dict[str, str]]]:
    normalized = deepcopy(spec)
    changes: list[dict[str, str]] = []

    title_before = str(normalized.get("title") or "")
    title_after = normalize_contextual_punctuation(title_before)
    if title_after != title_before:
        normalized["title"] = title_after
        changes.append({"location": "title", "before": title_before, "after": title_after})

    nodes = normalized.get("nodes") if isinstance(normalized.get("nodes"), list) else []
    for index, node in enumerate(nodes, 1):
        if not isinstance(node, dict):
            continue
        before = str(node.get("text") or "")
        after = normalize_contextual_punctuation(before)
        if after != before:
            node["text"] = after
            changes.append({"location": f"node:{index}", "before": before, "after": after})

    edges = normalized.get("edges") if isinstance(normalized.get("edges"), list) else []
    for index, edge in enumerate(edges, 1):
        if not isinstance(edge, dict) or not isinstance(edge.get("label"), str):
            continue
        before = str(edge.get("label") or "")
        after = normalize_contextual_punctuation(before)
        if after != before:
            edge["label"] = after
            changes.append({"location": f"edge:{index}", "before": before, "after": after})
    return normalized, changes


def sanitize_filename(value: str, fallback: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "-", value).strip("-.")
    return cleaned or fallback


def main() -> int:
    parser = argparse.ArgumentParser(description="Redraw structured flowchart JSON as SVG from semantic nodes/edges.")
    parser.add_argument("json_file", type=Path)
    parser.add_argument("--out-dir", type=Path, required=True)
    parser.add_argument("--allow-audit-errors", action="store_true", help="render even when structural audit reports errors")
    parser.add_argument("--engine", choices=("auto", "graphviz", "fallback"), default="auto", help="layout engine; auto prefers Graphviz dot")
    parser.add_argument("--allow-crossings", action="store_true", help="do not block on connector crossings found by geometric validation")
    args = parser.parse_args()

    try:
        raw = json.loads(args.json_file.read_text(encoding="utf-8"))
    except Exception as exc:
        parser.error(f"cannot parse JSON: {exc}")
    specs = raw if isinstance(raw, list) else [raw]
    if not all(isinstance(item, dict) for item in specs):
        parser.error("top-level JSON must be an object or array of objects")

    args.out_dir.mkdir(parents=True, exist_ok=True)
    manifest: list[dict[str, Any]] = []
    failed = False
    for idx, spec in enumerate(specs, 1):
        normalized_spec, punctuation_changes = normalize_spec_punctuation(spec)
        audit = audit_one(normalized_spec, idx) if audit_one is not None else None
        errors = int(audit["summary"]["severity"]["error"]) if audit else 0
        if errors and not args.allow_audit_errors:
            manifest.append({"index": idx, "id": normalized_spec.get("id"), "status": "blocked", "punctuation_changes": punctuation_changes, "audit": audit})
            failed = True
            continue
        try:
            svg, geometry, dot_source = render_svg(normalized_spec, engine=args.engine, allow_crossings=args.allow_crossings)
        except RuntimeError as exc:
            manifest.append({"index": idx, "id": normalized_spec.get("id"), "status": "blocked", "punctuation_changes": punctuation_changes, "audit": audit, "layout_error": str(exc)})
            failed = True
            continue
        stem = sanitize_filename(str(normalized_spec.get("id") or f"flowchart-{idx}"), f"flowchart-{idx}")
        out = args.out_dir / f"{stem}.svg"
        out.write_text(svg, encoding="utf-8")
        dot_out = None
        if dot_source is not None:
            dot_out = args.out_dir / f"{stem}.dot"
            dot_out.write_text(dot_source, encoding="utf-8")
        manifest.append({"index": idx, "id": normalized_spec.get("id"), "status": "rendered", "output": str(out), "dot": str(dot_out) if dot_out else None, "punctuation_changes": punctuation_changes, "audit": audit, "geometry": geometry})

    manifest_path = args.out_dir / "redraw-manifest.json"
    manifest_path.write_text(json.dumps({"flowcharts": manifest}, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"manifest": str(manifest_path), "flowcharts": manifest}, ensure_ascii=False, indent=2))
    return 2 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
