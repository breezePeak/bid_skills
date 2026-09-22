#!/usr/bin/env python3
"""Audit a structured directed flowchart before layout/rendering."""
from __future__ import annotations

import argparse
import json
from collections import defaultdict, deque
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any
import re
from punctuation_context import audit_fragment


@dataclass
class Issue:
    code: str
    severity: str
    location: str
    message: str
    suggestion: str | None = None



def contextual_punctuation_issues(text: str, location: str) -> list[Issue]:
    return [
        Issue(item.code, item.severity, location, f"{item.message} 原文：{text}",
              f"建议：{item.expected}" if item.expected else "按局部语境修正。")
        for item in audit_fragment(text)
    ]

def audit_one(spec: dict[str, Any], index: int = 1) -> dict[str, Any]:
    issues: list[Issue] = []
    prefix = f"flowchart:{index}"
    nodes = spec.get("nodes") if isinstance(spec.get("nodes"), list) else []
    edges = spec.get("edges") if isinstance(spec.get("edges"), list) else []
    direction = spec.get("direction")
    if direction not in (None, "TB", "LR"):
        issues.append(Issue("invalid-direction", "error", prefix, f"不支持的方向：{direction}", "使用 TB 或 LR。"))

    node_map: dict[str, dict[str, Any]] = {}
    for n_idx, node in enumerate(nodes, 1):
        loc = f"{prefix}/node:{n_idx}"
        if not isinstance(node, dict):
            issues.append(Issue("invalid-node", "error", loc, "节点不是对象。"))
            continue
        node_id = node.get("id")
        if not isinstance(node_id, str) or not node_id:
            issues.append(Issue("node-id-missing", "error", loc, "节点缺少稳定 id。"))
            continue
        if node_id in node_map:
            issues.append(Issue("duplicate-node-id", "error", loc, f"节点 id 重复：{node_id}"))
        node_map[node_id] = node
        text = node.get("text")
        if not isinstance(text, str) or not text.strip():
            issues.append(Issue("node-text-empty", "error", loc, f"节点 {node_id} 文本为空。"))
        else:
            issues.extend(contextual_punctuation_issues(text, loc))
            if len(text.strip()) > 48:
                issues.append(Issue("node-text-long", "warning", loc, f"节点 {node_id} 文本较长，布局前需要测量并换行。", "按实际字体宽度计算节点尺寸，不要固定高度。"))

    outgoing: dict[str, list[tuple[str, dict[str, Any]]]] = defaultdict(list)
    incoming: dict[str, list[tuple[str, dict[str, Any]]]] = defaultdict(list)
    edge_keys: set[tuple[str, str, str]] = set()
    for e_idx, edge in enumerate(edges, 1):
        loc = f"{prefix}/edge:{e_idx}"
        if not isinstance(edge, dict):
            issues.append(Issue("invalid-edge", "error", loc, "连线不是对象。"))
            continue
        src, dst = edge.get("from"), edge.get("to")
        label = edge.get("label") or ""
        if src not in node_map or dst not in node_map:
            issues.append(Issue("edge-endpoint-missing", "error", loc, f"连线端点不存在：{src} -> {dst}"))
            continue
        key = (str(src), str(dst), str(label))
        if key in edge_keys:
            issues.append(Issue("duplicate-edge", "warning", loc, f"重复连线：{src} -> {dst}"))
        edge_keys.add(key)
        outgoing[str(src)].append((str(dst), edge))
        incoming[str(dst)].append((str(src), edge))
        if isinstance(label, str) and label:
            issues.extend(contextual_punctuation_issues(label, loc))
        if src == dst:
            issues.append(Issue("self-loop", "warning", loc, f"节点 {src} 存在自环。", "确认该循环确有业务意义。"))

    starts = [nid for nid, node in node_map.items() if node.get("type") == "start"]
    ends = [nid for nid, node in node_map.items() if node.get("type") == "end"]
    if not starts:
        starts = [nid for nid in node_map if not incoming[nid]]
        if node_map:
            issues.append(Issue("start-node-not-explicit", "info", prefix, "未声明 start 节点，使用零入度节点做可达性检查。"))
    if not ends:
        ends = [nid for nid in node_map if not outgoing[nid]]
        if node_map:
            issues.append(Issue("end-node-not-explicit", "info", prefix, "未声明 end 节点，使用零出度节点做终点检查。"))

    for sid in starts:
        if incoming[sid]:
            issues.append(Issue("start-has-incoming-edge", "warning", f"{prefix}/node:{sid}", f"开始节点 {sid} 存在入边。", "确认箭头方向是否写反。"))
    for eid in ends:
        if outgoing[eid]:
            issues.append(Issue("end-has-outgoing-edge", "warning", f"{prefix}/node:{eid}", f"结束节点 {eid} 存在出边。", "确认箭头方向是否写反。"))

    reachable: set[str] = set()
    dq = deque(starts)
    while dq:
        nid = dq.popleft()
        if nid in reachable:
            continue
        reachable.add(nid)
        dq.extend(dst for dst, _ in outgoing[nid])
    for nid in node_map:
        if nid not in reachable:
            issues.append(Issue("node-unreachable-from-start", "error", f"{prefix}/node:{nid}", f"节点 {nid} 无法从开始节点到达。", "检查遗漏连线或错误箭头方向。"))

    can_reach_end: set[str] = set()
    dq = deque(ends)
    while dq:
        nid = dq.popleft()
        if nid in can_reach_end:
            continue
        can_reach_end.add(nid)
        dq.extend(src for src, _ in incoming[nid])
    for nid in node_map:
        if ends and nid not in can_reach_end:
            issues.append(Issue("node-cannot-reach-end", "warning", f"{prefix}/node:{nid}", f"节点 {nid} 不能到达任何结束节点。", "确认是否存在遗漏出口、错误方向或刻意循环。"))

    for nid, node in node_map.items():
        outs = outgoing[nid]
        if node.get("type") == "decision":
            if len(outs) < 2:
                issues.append(Issue("decision-missing-branches", "error", f"{prefix}/node:{nid}", f"判断节点 {nid} 少于两个分支。"))
            elif len(outs) == 2 and any(not str(edge.get("label") or "").strip() for _, edge in outs):
                issues.append(Issue("decision-branch-label-missing", "warning", f"{prefix}/node:{nid}", f"判断节点 {nid} 的两个分支未全部标注。", "为‘是／否’、‘通过／不通过’或业务条件添加明确中文标签。"))
            labels = [str(edge.get("label") or "").strip() for _, edge in outs if str(edge.get("label") or "").strip()]
            if len(labels) != len(set(labels)):
                issues.append(Issue("decision-branch-label-duplicate", "error", f"{prefix}/node:{nid}", f"判断节点 {nid} 存在重复分支标签。", "不同出边必须能从标签上明确区分。"))

    # Directed-cycle detection. Cycles may be valid, so this is a warning only.
    state: dict[str, int] = {nid: 0 for nid in node_map}
    cycle_nodes: set[str] = set()
    def dfs(nid: str, stack: list[str]) -> None:
        state[nid] = 1
        stack.append(nid)
        for dst, _ in outgoing[nid]:
            if state.get(dst, 0) == 0:
                dfs(dst, stack)
            elif state.get(dst) == 1:
                try:
                    pos = stack.index(dst)
                    cycle_nodes.update(stack[pos:])
                except ValueError:
                    cycle_nodes.add(dst)
        stack.pop()
        state[nid] = 2
    for nid in node_map:
        if state[nid] == 0:
            dfs(nid, [])
    if cycle_nodes:
        issues.append(Issue("directed-cycle", "warning", prefix, "流程图包含有向循环：" + ", ".join(sorted(cycle_nodes)), "循环可以合法，但布局和箭头方向必须在最终图中清晰可辨。"))

    severity = {"error": 0, "warning": 0, "info": 0}
    for issue in issues:
        severity[issue.severity] += 1
    return {
        "id": spec.get("id") or f"flowchart-{index}",
        "title": spec.get("title"),
        "summary": {"nodes": len(nodes), "edges": len(edges), "issues": len(issues), "severity": severity},
        "issues": [asdict(i) for i in issues],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit structured flowchart semantics before layout/rendering.")
    parser.add_argument("json_file", type=Path)
    parser.add_argument("--json-out", type=Path)
    args = parser.parse_args()
    try:
        raw = json.loads(args.json_file.read_text(encoding="utf-8"))
    except Exception as exc:
        parser.error(f"cannot parse JSON: {exc}")
    specs = raw if isinstance(raw, list) else [raw]
    if not all(isinstance(item, dict) for item in specs):
        parser.error("top-level JSON must be an object or array of objects")
    results = [audit_one(spec, idx) for idx, spec in enumerate(specs, 1)]
    total_errors = sum(r["summary"]["severity"]["error"] for r in results)
    result = {"flowcharts": results, "error_count": total_errors}
    encoded = json.dumps(result, ensure_ascii=False, indent=2)
    if args.json_out:
        args.json_out.parent.mkdir(parents=True, exist_ok=True)
        args.json_out.write_text(encoded + "\n", encoding="utf-8")
    print(encoded)
    return 2 if total_errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
