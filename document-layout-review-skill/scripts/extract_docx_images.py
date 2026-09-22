#!/usr/bin/env python3
"""Extract embedded DOCX images with placement context into a review directory."""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import zipfile
from pathlib import Path, PurePosixPath
from typing import Any
from xml.etree import ElementTree as ET

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "v": "urn:schemas-microsoft-com:vml",
}
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def safe_name(value: str, fallback: str) -> str:
    value = re.sub(r"[^A-Za-z0-9._-]+", "-", value).strip("-.")
    return value or fallback


def resolve_target(base: str, target: str) -> str:
    # Relationships are relative to word/document.xml's directory.
    if target.startswith("/"):
        return target.lstrip("/")
    path = PurePosixPath(base).parent / target
    parts: list[str] = []
    for part in path.parts:
        if part in ("", "."):
            continue
        if part == "..":
            if parts:
                parts.pop()
            continue
        parts.append(part)
    return "/".join(parts)


def paragraph_text(p: ET.Element) -> str:
    texts = [el.text or "" for el in p.findall(".//w:t", NS)]
    return "".join(texts).strip()


def image_dimensions(data: bytes) -> tuple[int | None, int | None, str | None]:
    if data.startswith(b"\x89PNG\r\n\x1a\n") and len(data) >= 24:
        return int.from_bytes(data[16:20], "big"), int.from_bytes(data[20:24], "big"), "png"
    if data.startswith(b"\xff\xd8"):
        i = 2
        while i + 9 < len(data):
            if data[i] != 0xFF:
                i += 1
                continue
            marker = data[i + 1]
            i += 2
            if marker in (0xD8, 0xD9):
                continue
            if i + 2 > len(data):
                break
            seg_len = int.from_bytes(data[i:i+2], "big")
            if seg_len < 2 or i + seg_len > len(data):
                break
            if marker in {0xC0,0xC1,0xC2,0xC3,0xC5,0xC6,0xC7,0xC9,0xCA,0xCB,0xCD,0xCE,0xCF} and seg_len >= 7:
                h = int.from_bytes(data[i+3:i+5], "big")
                w = int.from_bytes(data[i+5:i+7], "big")
                return w, h, "jpeg"
            i += seg_len
        return None, None, "jpeg"
    if data.startswith(b"GIF87a") or data.startswith(b"GIF89a"):
        if len(data) >= 10:
            return int.from_bytes(data[6:8], "little"), int.from_bytes(data[8:10], "little"), "gif"
        return None, None, "gif"
    return None, None, None


def main() -> int:
    parser = argparse.ArgumentParser(description="Extract embedded images from a DOCX for visual review.")
    parser.add_argument("docx", type=Path)
    parser.add_argument("--out-dir", type=Path, required=True)
    args = parser.parse_args()

    if not args.docx.is_file():
        parser.error("DOCX does not exist")
    args.out_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(args.docx) as zf:
        names = set(zf.namelist())
        if "word/document.xml" not in names or "word/_rels/document.xml.rels" not in names:
            parser.error("not a normal DOCX package")
        document = ET.fromstring(zf.read("word/document.xml"))
        rels_root = ET.fromstring(zf.read("word/_rels/document.xml.rels"))
        rels: dict[str, str] = {}
        for rel in rels_root.findall(f"{{{REL_NS}}}Relationship"):
            rid = rel.attrib.get("Id")
            target = rel.attrib.get("Target")
            mode = rel.attrib.get("TargetMode")
            if rid and target and mode != "External":
                rels[rid] = resolve_target("word/document.xml", target)

        placements: list[dict[str, Any]] = []
        written: dict[str, str] = {}
        paragraph_index = 0
        image_index = 0
        for p in document.findall(".//w:p", NS):
            paragraph_index += 1
            context = paragraph_text(p)
            refs: list[tuple[str, str, str, str]] = []
            for blip in p.findall(".//a:blip", NS):
                rid = blip.attrib.get(f"{{{NS['r']}}}embed")
                if rid:
                    refs.append((rid, "drawingml", "", ""))
            for imagedata in p.findall(".//v:imagedata", NS):
                rid = imagedata.attrib.get(f"{{{NS['r']}}}id")
                if rid:
                    refs.append((rid, "vml", "", ""))

            doc_prs = p.findall(".//wp:docPr", NS)
            alts = [(d.attrib.get("name") or "", d.attrib.get("title") or "", d.attrib.get("descr") or "") for d in doc_prs]
            for local_idx, (rid, kind, _, _) in enumerate(refs):
                target = rels.get(rid)
                if not target or target not in names:
                    continue
                data = zf.read(target)
                sha = hashlib.sha256(data).hexdigest()
                width, height, media_type = image_dimensions(data)
                ext = Path(target).suffix.lower() or (f".{media_type}" if media_type else ".bin")
                if sha not in written:
                    filename = safe_name(f"image-{len(written)+1:03d}-{sha[:10]}{ext}", f"image-{len(written)+1:03d}{ext}")
                    (args.out_dir / filename).write_bytes(data)
                    written[sha] = filename
                image_index += 1
                name, title, descr = alts[min(local_idx, len(alts)-1)] if alts else ("", "", "")
                placements.append({
                    "placement": image_index,
                    "paragraph": paragraph_index,
                    "relationship_id": rid,
                    "source_kind": kind,
                    "media_path": target,
                    "output": written[sha],
                    "sha256": sha,
                    "bytes": len(data),
                    "width": width,
                    "height": height,
                    "media_type": media_type,
                    "name": name,
                    "title": title,
                    "description": descr,
                    "paragraph_text": context,
                })

    manifest = {
        "source": str(args.docx),
        "unique_images": len(written),
        "placements": placements,
    }
    manifest_path = args.out_dir / "images-manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"manifest": str(manifest_path), **manifest}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
