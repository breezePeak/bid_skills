#!/usr/bin/env python3
"""Audit/repair punctuation by paragraph context while preserving DOCX run formatting.

DOCX repair edits only existing w:t text nodes. It never rebuilds paragraphs/runs.
If normalization changes paragraph text length, the repair is blocked for that paragraph
and reported for model/manual handling.
"""
from __future__ import annotations

import argparse
import json
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from punctuation_context import audit_fragment, normalize_fragment

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
def qn(name: str) -> str: return f"{{{W}}}{name}"


def paragraph_text_nodes(p: ET.Element):
    return p.findall(".//w:t", NS)


def paragraph_text(p: ET.Element) -> str:
    return "".join(t.text or "" for t in paragraph_text_nodes(p))


def process_xml(data: bytes, part: str, repair: bool):
    root = ET.fromstring(data)
    issues = []
    changes = []
    blocked = []
    for idx, p in enumerate(root.findall(".//w:p", NS), 1):
        nodes = paragraph_text_nodes(p)
        if not nodes:
            continue
        before = "".join(t.text or "" for t in nodes)
        if not before:
            continue
        for item in audit_fragment(before):
            issues.append({"part": part, "paragraph": idx, **item.__dict__})
        if not repair:
            continue
        after, kind = normalize_fragment(before)
        if after == before:
            continue
        if len(after) != len(before):
            blocked.append({"part": part, "paragraph": idx, "reason": "length-change", "context": kind, "before": before, "after": after})
            continue
        offset = 0
        changed_nodes = 0
        for t in nodes:
            old = t.text or ""
            new = after[offset:offset+len(old)]
            offset += len(old)
            if new != old:
                t.text = new
                changed_nodes += 1
        changes.append({"part": part, "paragraph": idx, "context": kind, "changed_text_nodes": changed_nodes, "before": before, "after": after})
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), issues, changes, blocked


def process_docx(src: Path, out: Path | None):
    issues, changes, blocked = [], [], []
    replacements = {}
    with zipfile.ZipFile(src, "r") as z:
        parts = [n for n in z.namelist() if n == "word/document.xml" or n.startswith("word/header") and n.endswith(".xml") or n.startswith("word/footer") and n.endswith(".xml")]
        for part in parts:
            updated, part_issues, part_changes, part_blocked = process_xml(z.read(part), part, out is not None)
            issues.extend(part_issues); changes.extend(part_changes); blocked.extend(part_blocked)
            if out is not None:
                replacements[part] = updated
        if out is not None:
            out.parent.mkdir(parents=True, exist_ok=True)
            with zipfile.ZipFile(out, "w") as dst:
                for info in z.infolist():
                    dst.writestr(info, replacements.get(info.filename, z.read(info.filename)))
    return {"input": str(src), "output": str(out) if out else None, "issue_count": len(issues), "issues": issues, "change_count": len(changes), "changes": changes, "blocked_count": len(blocked), "blocked": blocked}


def process_text(src: Path, out: Path | None):
    raw = src.read_text(encoding="utf-8")
    issues, changes = [], []
    rendered = []
    for i, line in enumerate(raw.splitlines(keepends=True), 1):
        ending = "\n" if line.endswith("\n") else ""
        body = line[:-1] if ending else line
        for item in audit_fragment(body):
            issues.append({"line": i, **item.__dict__})
        after, kind = normalize_fragment(body)
        if after != body:
            changes.append({"line": i, "context": kind, "before": body, "after": after})
        rendered.append(after + ending)
    if out is not None:
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text("".join(rendered), encoding="utf-8")
    return {"input": str(src), "output": str(out) if out else None, "issue_count": len(issues), "issues": issues, "change_count": len(changes), "changes": changes, "blocked_count": 0, "blocked": []}


def main() -> int:
    ap = argparse.ArgumentParser(description="Context-aware punctuation: Chinese full-width; English/code/formulas half-width, preserving DOCX runs.")
    ap.add_argument("input", type=Path)
    ap.add_argument("--out", type=Path)
    ap.add_argument("--json-out", type=Path)
    args = ap.parse_args()
    if not args.input.is_file(): ap.error("input does not exist")
    result = process_docx(args.input, args.out) if args.input.suffix.lower() == ".docx" else process_text(args.input, args.out)
    text = json.dumps(result, ensure_ascii=False, indent=2)
    if args.json_out:
        args.json_out.parent.mkdir(parents=True, exist_ok=True)
        args.json_out.write_text(text + "\n", encoding="utf-8")
    print(text)
    return 2 if result["issue_count"] and args.out is None else (3 if result["blocked_count"] else 0)

if __name__ == "__main__":
    raise SystemExit(main())
