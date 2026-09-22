#!/usr/bin/env python3
"""Conservative useless-whitespace repair that preserves DOCX run structure.

The decision is made on the complete visible paragraph text, so whitespace split
across run boundaries is still detected. Characters are removed or normalized in
their original w:t nodes; runs, fonts, bold, size and paragraph properties are not
rebuilt.
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

# Chinese date/time and quantity units should be attached to their numbers in normal
# Chinese prose. Keep this list deliberately semantic instead of deleting every
# Chinese/Latin boundary space, so legitimate text such as "中文 API" remains intact.
DATE_TIME_UNITS = ("年", "月", "日", "号", "时", "分", "秒")
COUNT_UNITS = (
    "平方公里", "平方米", "立方米", "公里", "厘米", "毫米", "万元", "亿元",
    "套", "份", "台", "个", "项", "次", "人", "家", "组", "类", "页", "章",
    "节", "条", "款", "批", "处", "点", "座", "册", "张", "幅", "米", "元",
)
NUMBER_SUFFIXES = tuple(sorted(set(DATE_TIME_UNITS + COUNT_UNITS + ("%", "％", "℃")), key=len, reverse=True))
NUMBER_SUFFIX_RE = "(?:" + "|".join(re.escape(x) for x in NUMBER_SUFFIXES) + ")"
COUNT_UNIT_RE = "(?:" + "|".join(re.escape(x) for x in sorted(COUNT_UNITS, key=len, reverse=True)) + ")"
NUMBER_RE = r"\d+(?:[.,]\d+)?"


def text_nodes(p: ET.Element) -> list[ET.Element]:
    return p.findall(".//w:t", NS)


def _remove_match_spaces(s: str, pattern: str, rm: set[int]) -> None:
    for m in re.finditer(pattern, s):
        for i in range(m.start(), m.end()):
            if s[i] in SPACE_CHARS:
                rm.add(i)


def whitespace_edits(s: str) -> tuple[set[int], dict[int, str]]:
    # Only enable these rules in Chinese/mixed prose to avoid damaging code,
    # commands, URLs and pure English text.
    if context_kind(s) not in {"chinese", "mixed"} or not re.search(fr"[{HAN}]", s):
        return set(), {}

    rm: set[int] = set()
    replacements: dict[int, str] = {}

    # Zero-width spaces/BOM are invisible dirty characters in normal prose.
    for m in ZERO_WIDTH_RE.finditer(s):
        rm.update(range(m.start(), m.end()))

    # Leading/trailing spaces in paragraphs or cells.
    for m in re.finditer(fr"^{SP}+|{SP}+$", s):
        rm.update(range(m.start(), m.end()))

    # Spaces between adjacent Chinese characters.
    for m in re.finditer(fr"(?<=[{HAN}]){SP}+(?=[{HAN}])", s):
        rm.update(range(m.start(), m.end()))

    # Spaces around Chinese punctuation, brackets and quotation marks.
    for m in re.finditer(fr"{SP}+(?=[{re.escape(CLOSE)}])", s):
        rm.update(range(m.start(), m.end()))
    for m in re.finditer(fr"(?<=[{re.escape(OPEN)}]){SP}+", s):
        rm.update(range(m.start(), m.end()))
    for m in re.finditer(fr"(?<=[{re.escape(CLOSE)}]){SP}+(?=\S)", s):
        rm.update(range(m.start(), m.end()))
    for m in re.finditer(fr"(?<=\S){SP}+(?=[{re.escape(OPEN)}])", s):
        rm.update(range(m.start(), m.end()))

    # Numeric date/time and Chinese-unit spacing:
    #   2026 年 10 月 31 日 -> 2026年10月31日
    #   1 套 / 2 份         -> 1套 / 2份
    _remove_match_spaces(s, fr"(?<=\d){SP}+(?={NUMBER_SUFFIX_RE})", rm)
    _remove_match_spaces(s, fr"(?<=[年月日号时分秒]){SP}+(?=\d)", rm)

    # Remove the space before a Chinese quantity only when the following number has
    # a known Chinese counter/unit. This fixes "软件 1 套" -> "软件1套" without
    # globally removing legitimate Chinese/English spacing.
    _remove_match_spaces(
        s,
        fr"(?<=[{HAN}]){SP}+(?={NUMBER_RE}{SP}*{COUNT_UNIT_RE})",
        rm,
    )

    # Normalize remaining full-width/NBSP spaces to ASCII space. This keeps legal
    # mixed-language single spaces while removing hidden-space variants.
    for i, ch in enumerate(s):
        if ch in {"\u3000", "\u00a0"}:
            replacements[i] = " "

    # Multiple spaces next to Chinese text are usually paste pollution. Keep one to
    # avoid damaging legitimate "中文 API" / "API 中文" single spaces.
    cjk_side = HAN + re.escape(OPEN + CLOSE)
    for m in re.finditer(
        fr"(?<=[{cjk_side}]){SP}{{2,}}(?=\S)|(?<=\S){SP}{{2,}}(?=[{cjk_side}])",
        s,
    ):
        rm.update(range(m.start() + 1, m.end()))

    # Deletion wins over replacement.
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
