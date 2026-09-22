#!/usr/bin/env python3
"""DOCX object-level mutation guard for document-layout-review.

The guard records semantic/layout signatures from a baseline DOCX and verifies
that a candidate changed only what the declared repair scope is allowed to
change.

It intentionally does not decide whether the target repair is correct. It only
prevents collateral changes from being accepted as a successful repair.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def q(ns: str, local: str) -> str:
    return f"{{{ns}}}{local}"


def local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def stable_elem(elem: ET.Element, drop_descendants: set[str] | None = None) -> str:
    """Serialize XML semantically with sorted attributes.

    Namespace prefixes and attribute ordering are ignored. Optional descendant
    local-names can be omitted (used to exclude sectPr from paragraph style).
    """
    drop_descendants = drop_descendants or set()

    def walk(e: ET.Element) -> str:
        lname = local(e.tag)
        if lname in drop_descendants:
            return ""
        attrs = sorted((local(k), v) for k, v in e.attrib.items())
        head = "<" + lname
        for k, v in attrs:
            head += f" {k}={json.dumps(v, ensure_ascii=False)}"
        head += ">"
        text = e.text or ""
        children = "".join(walk(c) for c in list(e))
        tail = e.tail or ""
        return head + text + children + f"</{lname}>" + tail

    return walk(elem)


def story_part(name: str) -> bool:
    if name == "word/document.xml":
        return True
    base = Path(name).name
    return base.startswith(("header", "footer")) and base.endswith(".xml") or base in {
        "footnotes.xml",
        "endnotes.xml",
        "comments.xml",
    }


def parse_story_parts(z: zipfile.ZipFile) -> list[tuple[str, ET.Element]]:
    out: list[tuple[str, ET.Element]] = []
    for name in sorted(z.namelist()):
        if not story_part(name):
            continue
        try:
            out.append((name, ET.fromstring(z.read(name))))
        except ET.ParseError:
            continue
    return out


def collect_text_content(parts: list[tuple[str, ET.Element]]) -> str:
    rows: list[tuple[str, str, str]] = []
    for name, root in parts:
        for e in root.iter():
            lname = local(e.tag)
            if lname in {"t", "instrText", "delText"}:
                rows.append((name, lname, e.text or ""))
    return sha(json.dumps(rows, ensure_ascii=False, separators=(",", ":")).encode("utf-8"))


def collect_styles(parts: list[tuple[str, ET.Element]], outside_tables_only: bool = False) -> str:
    rows: list[tuple[str, str]] = []

    def walk(name: str, e: ET.Element, inside_tbl: bool) -> None:
        here_tbl = inside_tbl or e.tag == q(W, "tbl")
        if e.tag in {q(W, "rPr"), q(W, "pPr")}:
            if not outside_tables_only or not here_tbl:
                rows.append((name, stable_elem(e, drop_descendants={"sectPr"})))
        for c in list(e):
            walk(name, c, here_tbl)

    for name, root in parts:
        walk(name, root, False)
    return sha(json.dumps(rows, ensure_ascii=False, separators=(",", ":")).encode("utf-8"))


def table_structure(parts: list[tuple[str, ET.Element]]) -> str:
    tables: list[dict] = []
    for name, root in parts:
        for ti, tbl in enumerate(root.iter(q(W, "tbl"))):
            rows = []
            for tr in [c for c in list(tbl) if c.tag == q(W, "tr")]:
                cells = []
                for tc in [c for c in list(tr) if c.tag == q(W, "tc")]:
                    tcpr = tc.find(q(W, "tcPr"))
                    span = None
                    merge = None
                    if tcpr is not None:
                        gs = tcpr.find(q(W, "gridSpan"))
                        vm = tcpr.find(q(W, "vMerge"))
                        if gs is not None:
                            span = gs.attrib.get(q(W, "val")) or gs.attrib.get("val")
                        if vm is not None:
                            merge = vm.attrib.get(q(W, "val")) or vm.attrib.get("val") or "continue"
                    cells.append({"gridSpan": span, "vMerge": merge})
                rows.append(cells)
            tables.append({"part": name, "index": ti, "rows": rows})
    return sha(json.dumps(tables, ensure_ascii=False, separators=(",", ":")).encode("utf-8"))


def table_geometry(parts: list[tuple[str, ET.Element]]) -> str:
    rows: list[tuple[str, int, str]] = []
    for name, root in parts:
        for ti, tbl in enumerate(root.iter(q(W, "tbl"))):
            chunks: list[str] = []
            for e in tbl.iter():
                if e.tag in {q(W, "tblPr"), q(W, "tblGrid"), q(W, "trPr"), q(W, "tcPr")}:
                    chunks.append(stable_elem(e))
            rows.append((name, ti, "".join(chunks)))
    return sha(json.dumps(rows, ensure_ascii=False, separators=(",", ":")).encode("utf-8"))


def drawings(parts: list[tuple[str, ET.Element]]) -> str:
    rows: list[tuple[str, str]] = []
    targets = {q(W, "drawing"), q(W, "pict")}
    for name, root in parts:
        for e in root.iter():
            if e.tag in targets:
                rows.append((name, stable_elem(e)))
    return sha(json.dumps(rows, ensure_ascii=False, separators=(",", ":")).encode("utf-8"))


def sections(parts: list[tuple[str, ET.Element]]) -> str:
    rows: list[tuple[str, str]] = []
    for name, root in parts:
        for e in root.iter(q(W, "sectPr")):
            rows.append((name, stable_elem(e)))
    return sha(json.dumps(rows, ensure_ascii=False, separators=(",", ":")).encode("utf-8"))


def block_structure(parts: list[tuple[str, ET.Element]]) -> str:
    keep = {"body", "hdr", "ftr", "footnote", "endnote", "comment", "p", "tbl", "tr", "tc", "sdt", "txbxContent"}
    rows: list[tuple[str, tuple[str, ...]]] = []
    for name, root in parts:
        sig: list[str] = []
        for e in root.iter():
            lname = local(e.tag)
            if lname in keep:
                sig.append(lname)
        rows.append((name, tuple(sig)))
    return sha(json.dumps(rows, ensure_ascii=False, separators=(",", ":")).encode("utf-8"))


def package_hash_group(z: zipfile.ZipFile, predicate) -> str:
    rows: list[tuple[str, str]] = []
    for name in sorted(z.namelist()):
        if predicate(name):
            rows.append((name, sha(z.read(name))))
    return sha(json.dumps(rows, separators=(",", ":")).encode("utf-8"))


def snapshot_docx(path: Path) -> dict:
    with zipfile.ZipFile(path, "r") as z:
        parts = parse_story_parts(z)
        return {
            "source": str(path),
            "source_sha256": sha(path.read_bytes()),
            "text_content": collect_text_content(parts),
            "text_style": collect_styles(parts, outside_tables_only=False),
            "non_table_text_style": collect_styles(parts, outside_tables_only=True),
            "block_structure": block_structure(parts),
            "table_structure": table_structure(parts),
            "table_geometry": table_geometry(parts),
            "drawings": drawings(parts),
            "sections": sections(parts),
            "media": package_hash_group(z, lambda n: n.startswith("word/media/")),
            "relationships": package_hash_group(z, lambda n: n.startswith("word/") and n.endswith(".rels")),
        }


INVARIANTS = {
    "text-style": {
        "text_content", "block_structure", "table_structure", "table_geometry",
        "drawings", "sections", "media", "relationships",
    },
    "text-content": {
        "text_style", "block_structure", "table_structure", "table_geometry",
        "drawings", "sections", "media", "relationships",
    },
    "table-text-style": {
        "text_content", "non_table_text_style", "block_structure", "table_structure",
        "table_geometry", "drawings", "sections", "media", "relationships",
    },
    "table-layout": {
        "text_content", "text_style", "block_structure", "table_structure",
        "drawings", "sections", "media", "relationships",
    },
    "image-layout": {
        "text_content", "text_style", "block_structure", "table_structure",
        "table_geometry", "sections",
    },
    "flowchart-redraw": {
        "text_content", "text_style", "block_structure", "table_structure",
        "table_geometry", "sections",
    },
    "page-layout": {
        "text_content", "text_style", "block_structure", "table_structure",
        "table_geometry", "drawings", "media", "relationships",
    },
}


def cmd_snapshot(args: argparse.Namespace) -> int:
    src = Path(args.docx)
    data = snapshot_docx(src)
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"PASS snapshot -> {out}")
    return 0


def cmd_compare(args: argparse.Namespace) -> int:
    if args.scope not in INVARIANTS:
        print(f"ERROR unsupported scope: {args.scope}", file=sys.stderr)
        return 2
    baseline = json.loads(Path(args.guard).read_text(encoding="utf-8"))
    candidate = snapshot_docx(Path(args.docx))
    changed = [k for k in sorted(INVARIANTS[args.scope]) if baseline.get(k) != candidate.get(k)]
    if changed:
        print(f"FAIL scope={args.scope}")
        for k in changed:
            print(f"  changed protected invariant: {k}")
        return 1
    print(f"PASS scope={args.scope}; protected invariants unchanged")
    return 0


def main() -> int:
    p = argparse.ArgumentParser(description="Guard DOCX non-target objects against collateral edits")
    sub = p.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("snapshot", help="create baseline signatures")
    s.add_argument("docx")
    s.add_argument("--out", required=True)
    s.set_defaults(func=cmd_snapshot)

    c = sub.add_parser("compare", help="compare candidate against baseline for a repair scope")
    c.add_argument("guard")
    c.add_argument("docx")
    c.add_argument("--scope", required=True, choices=sorted(INVARIANTS))
    c.set_defaults(func=cmd_compare)

    args = p.parse_args()
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
