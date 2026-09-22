#!/usr/bin/env python3
"""Conservative useless-whitespace repair that preserves DOCX run structure.

The decision is made on the complete visible paragraph text, so whitespace split
across run boundaries is still detected. Characters are removed or normalized in their
original w:t nodes; runs, fonts, bold, size and paragraph properties are not rebuilt.
"""
from __future__ import annotations

import argparse
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from punctuation_context import context_kind

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
HAN = r"\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff"
CLOSE = "，。；：！？、）】》」』”’"
OPEN = "（【《「『“‘"
SPACE_CHARS = " \u3000\u00a0"
SP = f"[{re.escape(SPACE_CHARS)}]"
ZERO_WIDTH_RE = re.compile(r"[\u200b\ufeff]")


def text_nodes(p: ET.Element) -> list[ET.Element]:
    return p.findall(".//w:t", NS)


def whitespace_edits(s: str) -> tuple[set[int], dict[int, str]]:
    # 只对明确的中文/中英混排正文启用，避免破坏代码、命令和纯英文。
    if context_kind(s) not in {"chinese", "mixed"} or not re.search(fr"[{HAN}]", s):
        return set(), {}

    rm: set[int] = set()
    replacements: dict[int, str] = {}

    # 零宽空格/BOM 在普通中文正文中属于不可见脏字符。
    for m in ZERO_WIDTH_RE.finditer(s):
        rm.update(range(m.start(), m.end()))

    # 段首/段尾或单元格首尾的普通、全角、NBSP 空格。
    for m in re.finditer(fr"^{SP}+|{SP}+$", s):
        rm.update(range(m.start(), m.end()))

    # 中文字符之间不应存在空格。
    for m in re.finditer(fr"(?<=[{HAN}]){SP}+(?=[{HAN}])", s):
        rm.update(range(m.start(), m.end()))

    # 中文标点、括号、引号周边的无意义空格。
    for m in re.finditer(fr"{SP}+(?=[{re.escape(CLOSE)}])", s):
        rm.update(range(m.start(), m.end()))
    for m in re.finditer(fr"(?<=[{re.escape(OPEN)}]){SP}+", s):
        rm.update(range(m.start(), m.end()))
    for m in re.finditer(fr"(?<=[{re.escape(CLOSE)}]){SP}+(?=\S)", s):
        rm.update(range(m.start(), m.end()))
    for m in re.finditer(fr"(?<=\S){SP}+(?=[{re.escape(OPEN)}])", s):
        rm.update(range(m.start(), m.end()))

    # 未被删除的全角空格/NBSP 统一成普通 ASCII 空格。这样既能清理隐藏字符，
    # 又不会把合法的“中文 API”式中英单空格直接吃掉。
    for i, ch in enumerate(s):
        if ch in {"\u3000", "\u00a0"}:
            replacements[i] = " "

    # 中英混排中连续两个以上空格通常是粘贴污染；保留一个，避免误伤合法的
    # “中文 API” / “API 中文”单空格。
    cjk_side = HAN + re.escape(OPEN + CLOSE)
    for m in re.finditer(
        fr"(?<=[{cjk_side}]){SP}{{2,}}(?=\S)|(?<=\S){SP}{{2,}}(?=[{cjk_side}])",
        s,
    ):
        rm.update(range(m.start() + 1, m.end()))

    # 删除优先于替换。
    for i in rm:
        replacements.pop(i, None)
    return rm, replacements


def process_xml(data: bytes) -> tuple[bytes, list[dict]]:
    root = ET.fromstring(data)
    changes = []
    for idx, p in enumerate(root.findall(".//w:p", NS), 1):
        nodes = text_nodes(p)
        before = "".join(t.text or "" for t in nodes)
        if not before:
            continue
        rm, replacements = whitespace_edits(before)
        if not rm and not replacements:
            continue

        offset = 0
        changed_nodes = 0
        for t in nodes:
            old = t.text or ""
            chars = []
            for j, ch in enumerate(old):
                pos = offset + j
                if pos in rm:
                    continue
                chars.append(replacements.get(pos, ch))
            new = "".join(chars)
            offset += len(old)
            if new != old:
                t.text = new
                changed_nodes += 1

        after = "".join(t.text or "" for t in nodes)
        changes.append(
            {
                "paragraph": idx,
                "before": before,
                "after": after,
                "changed_text_nodes": changed_nodes,
            }
        )

    return ET.tostring(root, encoding="utf-8", xml_declaration=True), changes


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("input", type=Path)
    ap.add_argument("--out", type=Path, required=True)
    ap.add_argument("--json-out", type=Path)
    a = ap.parse_args()

    replacements: dict[str, bytes] = {}
    changes: list[dict] = []
    with zipfile.ZipFile(a.input) as z:
        parts = [
            n
            for n in z.namelist()
            if n == "word/document.xml"
            or (n.startswith("word/header") and n.endswith(".xml"))
            or (n.startswith("word/footer") and n.endswith(".xml"))
            or n in {"word/footnotes.xml", "word/endnotes.xml"}
        ]
        for part in parts:
            xml, part_changes = process_xml(z.read(part))
            replacements[part] = xml
            changes += [{"part": part, **x} for x in part_changes]

        a.out.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(a.out, "w") as dst:
            for info in z.infolist():
                dst.writestr(info, replacements.get(info.filename, z.read(info.filename)))

    result = {
        "input": str(a.input),
        "output": str(a.out),
        "change_count": len(changes),
        "changes": changes,
    }
    payload = json.dumps(result, ensure_ascii=False, indent=2)
    if a.json_out:
        a.json_out.parent.mkdir(parents=True, exist_ok=True)
        a.json_out.write_text(payload + "\n", encoding="utf-8")
    print(payload)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
