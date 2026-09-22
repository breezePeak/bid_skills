#!/usr/bin/env python3
"""Hard release-gate audit for punctuation/whitespace in visible DOCX text.

Scans document, headers, footers, footnotes/endnotes and text boxes (via w:t).
This complements contextual_punctuation.py by rejecting characters that can slip
through span heuristics, especially full-width ASCII quotation marks and hidden
whitespace defects.
"""
from __future__ import annotations

import argparse
import json
import re
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

from punctuation_context import audit_fragment, context_kind

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
HAN = r"\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff"
HAN_RE = re.compile(fr"[{HAN}]")
CJK_CLOSE = "，。；：！？、）】》」』”’"
CJK_OPEN = "（【《「『“‘"
FULLWIDTH_ASCII_QUOTES = {"＂", "＇"}
SPACE_CHARS = " \u3000\u00a0"
SP = f"[{re.escape(SPACE_CHARS)}]"


@dataclass
class Issue:
    severity: str
    code: str
    part: str
    paragraph: int
    message: str
    evidence: str


def text_of(p: ET.Element) -> str:
    return "".join(t.text or "" for t in p.findall(".//w:t", NS))


def excerpt(s: str, n: int = 180) -> str:
    s = s.replace("\n", "\\n")
    return s if len(s) <= n else s[: n - 1] + "…"


def audit_paragraph(text: str, part: str, idx: int) -> list[Issue]:
    if not text:
        return []
    out: list[Issue] = []
    kind = context_kind(text)

    # First run the shared contextual normalizer/auditor.
    for item in audit_fragment(text):
        out.append(Issue(item.severity, item.code, part, idx, item.message, excerpt(text)))

    # Full-width ASCII quotes are never acceptable as final punctuation.
    if any(ch in text for ch in FULLWIDTH_ASCII_QUOTES):
        out.append(
            Issue(
                "error",
                "fullwidth-ascii-quote",
                part,
                idx,
                "存在全角 ASCII 引号（＂/＇）。中文引号应使用“”/‘’，英文、代码和公式应使用半角引号。",
                excerpt(text),
            )
        )

    # Chinese prose: reject ASCII quotes when they quote Chinese or sit near Han text.
    if kind in {"chinese", "mixed"} and HAN_RE.search(text):
        if re.search(fr'["＂][^"＂\n]*[{HAN}][^"＂\n]*["＂]', text):
            out.append(
                Issue(
                    "error",
                    "ascii-double-quote-in-chinese",
                    part,
                    idx,
                    "中文语境中引用中文内容时必须使用中文双引号“”。",
                    excerpt(text),
                )
            )
        if re.search(fr"['＇][^'＇\n]*[{HAN}][^'＇\n]*['＇]", text):
            out.append(
                Issue(
                    "error",
                    "ascii-single-quote-in-chinese",
                    part,
                    idx,
                    "中文语境中引用中文内容时必须使用中文单引号‘’。",
                    excerpt(text),
                )
            )

        # Even an unmatched quote near Han is a blocker; it cannot be waved through by heuristics.
        for m in re.finditer(r'["＂\'＇]', text):
            left = text[max(0, m.start() - 12) : m.start()]
            right = text[m.end() : m.end() + 12]
            if HAN_RE.search(left + right):
                out.append(
                    Issue(
                        "error",
                        "ambiguous-ascii-quote-near-chinese",
                        part,
                        idx,
                        "中文正文附近存在 ASCII/全角 ASCII 引号，必须人工或规则确定为中文引号后再交付。",
                        excerpt(text),
                    )
                )
                break

        # Character-level sentence punctuation checks in clear Chinese adjacency.
        patterns = [
            (fr"(?<=[{HAN}]),|,(?=[{HAN}])", "ascii-comma-in-chinese", "中文语境使用中文逗号“，”"),
            (fr"(?<=[{HAN}]);|;(?=[{HAN}])", "ascii-semicolon-in-chinese", "中文语境使用中文分号“；”"),
            (fr"(?<=[{HAN}）】》”’])\?", "ascii-question-in-chinese", "中文语境使用中文问号“？”"),
            (fr"(?<=[{HAN}）】》”’])!", "ascii-exclamation-in-chinese", "中文语境使用中文叹号“！”"),
        ]
        for pat, code, msg in patterns:
            if re.search(pat, text):
                out.append(Issue("error", code, part, idx, msg, excerpt(text)))

        # Whitespace hard gates. All checks operate on concatenated visible text, so
        # spaces at run boundaries are included automatically.
        if re.search(fr"(?<=[{HAN}]){SP}+(?=[{HAN}])", text):
            out.append(
                Issue(
                    "error",
                    "space-between-han",
                    part,
                    idx,
                    "普通中文正文中相邻中文字符之间存在无用空格。",
                    excerpt(text),
                )
            )
        if re.search(fr"{SP}+(?=[{re.escape(CJK_CLOSE)}])", text):
            out.append(
                Issue(
                    "error",
                    "space-before-cjk-punctuation",
                    part,
                    idx,
                    "中文标点或右括号/右引号前存在无用空格。",
                    excerpt(text),
                )
            )
        if re.search(fr"(?<=[{re.escape(CJK_OPEN)}]){SP}+", text):
            out.append(
                Issue(
                    "error",
                    "space-after-cjk-opening-punctuation",
                    part,
                    idx,
                    "中文左括号/左引号后存在无用空格。",
                    excerpt(text),
                )
            )
        if re.search(fr"(?<=[{re.escape(CJK_CLOSE)}]){SP}+(?=\S)", text):
            out.append(
                Issue(
                    "error",
                    "space-after-cjk-punctuation",
                    part,
                    idx,
                    "中文标点或右括号/右引号后存在无用空格。",
                    excerpt(text),
                )
            )
        if re.search(fr"(?<=\S){SP}+(?=[{re.escape(CJK_OPEN)}])", text):
            out.append(
                Issue(
                    "error",
                    "space-before-cjk-opening-punctuation",
                    part,
                    idx,
                    "中文左括号/左引号前存在无用空格。",
                    excerpt(text),
                )
            )
        if re.search(fr"^{SP}+|{SP}+$", text):
            out.append(
                Issue(
                    "error",
                    "leading-or-trailing-whitespace",
                    part,
                    idx,
                    "中文正文或单元格首尾存在无用空格。",
                    excerpt(text),
                )
            )
        if re.search(
            fr"(?<=[{HAN}])[\u3000\u00a0]+(?=\S)|(?<=\S)[\u3000\u00a0]+(?=[{HAN}])",
            text,
        ):
            out.append(
                Issue(
                    "error",
                    "cjk-wide-or-nbsp",
                    part,
                    idx,
                    "中文正文附近存在全角空格或不换行空格（NBSP）。",
                    excerpt(text),
                )
            )
        if re.search(r"[\u200b\ufeff]", text):
            out.append(
                Issue(
                    "error",
                    "zero-width-whitespace",
                    part,
                    idx,
                    "正文中存在零宽空格/BOM 等不可见脏字符。",
                    excerpt(text),
                )
            )

    if re.search(r"(?:，，|。。|；；|：：|？？|！！|、、)", text):
        out.append(
            Issue(
                "error",
                "duplicated-chinese-punctuation",
                part,
                idx,
                "存在重复中文标点。",
                excerpt(text),
            )
        )
    return out


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("docx", type=Path)
    ap.add_argument("--json-out", type=Path)
    args = ap.parse_args()
    if not args.docx.is_file():
        ap.error("DOCX not found")

    issues: list[Issue] = []
    with zipfile.ZipFile(args.docx) as z:
        names = [
            n
            for n in z.namelist()
            if n == "word/document.xml"
            or (n.startswith("word/header") and n.endswith(".xml"))
            or (n.startswith("word/footer") and n.endswith(".xml"))
            or n in {"word/footnotes.xml", "word/endnotes.xml"}
        ]
        for name in names:
            try:
                root = ET.fromstring(z.read(name))
            except Exception:
                continue
            for idx, p in enumerate(root.findall(".//w:p", NS), 1):
                issues.extend(audit_paragraph(text_of(p), name, idx))

    # De-duplicate same code/part/paragraph caused by shared + hard patterns.
    uniq = []
    seen = set()
    for issue in issues:
        key = (issue.code, issue.part, issue.paragraph)
        if key not in seen:
            seen.add(key)
            uniq.append(issue)

    result = {
        "input": str(args.docx),
        "status": "passed" if not uniq else "failed",
        "issue_count": len(uniq),
        "issues": [asdict(i) for i in uniq],
    }
    payload = json.dumps(result, ensure_ascii=False, indent=2)
    if args.json_out:
        args.json_out.parent.mkdir(parents=True, exist_ok=True)
        args.json_out.write_text(payload + "\n", encoding="utf-8")
    print(payload)
    return 0 if not uniq else 2


if __name__ == "__main__":
    raise SystemExit(main())
