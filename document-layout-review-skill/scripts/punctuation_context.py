#!/usr/bin/env python3
"""Context-aware punctuation rules shared by document-layout-review.

Rules:
- Chinese prose -> full-width Chinese punctuation.
- English prose, code, paths, URLs, identifiers and formulas -> half-width punctuation.
- Mixed text -> decide from the local span; do not apply document-wide replacement.
"""
from __future__ import annotations
from dataclasses import dataclass
import re

HAN = r"\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff"
HAN_RE = re.compile(fr"[{HAN}]")
LATIN_RE = re.compile(r"[A-Za-z]")

CJK_TO_ASCII = {
    "，": ",", "。": ".", "；": ";", "：": ":", "？": "?", "！": "!",
    "（": "(", "）": ")", "／": "/", "“": '"', "”": '"', "‘": "'", "’": "'", "、": ",",
}
FULLWIDTH_ASCII_TO_ASCII = {
    "％": "%", "＋": "+", "－": "-", "＝": "=", "＜": "<", "＞": ">",
    "＃": "#", "＠": "@", "＆": "&", "＊": "*", "＿": "_", "｜": "|",
    "～": "~", "＾": "^", "＄": "$", "＼": "\\", "＂": '"', "＇": "'",
}
CODE_OR_TECH_RE = re.compile(
    r"(?:https?://|mailto:|[A-Za-z]:\\|(?:^|\s)(?:npm|pnpm|python|python3|node|git|curl|docker|kubectl)\s|"
    r"\b[A-Za-z_][A-Za-z0-9_]*\s*(?:==|!=|<=|>=|=>|:=|=|\+|-|\*|/|\^|%)|"
    r"(?:==|!=|<=|>=|=>|::|&&|\|\||\+\+|--))"
)
FORMULA_ALLOWED_RE = re.compile(
    r'^[A-Za-z0-9_\s.,;:!?()\[\]{}+\-*/=<>^%|~#@$\\'
    r'，。；：？！（）／％＋－＝＜＞＊＾＿｜～＃＠＆＄＼]+$'
)
FORMULA_OPERATOR_RE = re.compile(r"[=+\-*/<>^%]|[＝＋－＊／＜＞＾％]")
EMBEDDED_TECH_SPAN_RE = re.compile(
    r"[A-Za-z0-9_][A-Za-z0-9_\s（）／％＋－＝＜＞＊＾＿｜～＃＠＆＄＼"
    r"()\[\]{}+\-*/=<>^%|~#@$\\.,;:!?]*[A-Za-z0-9_）)]"
)

@dataclass(frozen=True)
class PunctuationIssue:
    code: str
    severity: str
    message: str
    token: str
    expected: str | None = None


def context_kind(text: str) -> str:
    s = text.strip()
    if not s:
        return "neutral"
    han = len(HAN_RE.findall(s))
    latin = len(LATIN_RE.findall(s))
    if han == 0 and CODE_OR_TECH_RE.search(s):
        return "technical"
    if han == 0 and FORMULA_ALLOWED_RE.fullmatch(s) and FORMULA_OPERATOR_RE.search(s):
        return "technical"
    if han == 0 and latin > 0:
        return "english"
    if han > 0 and (latin == 0 or han >= latin):
        return "chinese"
    return "mixed"


def _normalize_embedded_nonchinese(text: str) -> str:
    """Normalize clearly non-Chinese spans embedded in Chinese prose."""
    text = re.sub(r'（([^（）\n]*[A-Za-z0-9][^（）\n]*)）',
                  lambda m: f"({m.group(1)})" if not HAN_RE.search(m.group(1)) else m.group(0), text)
    text = re.sub(r'“([^“”\n]*[A-Za-z0-9][^“”\n]*)”',
                  lambda m: f'"{m.group(1)}"' if not HAN_RE.search(m.group(1)) else m.group(0), text)
    text = re.sub(r'‘([^‘’\n]*[A-Za-z0-9][^‘’\n]*)’',
                  lambda m: f"'{m.group(1)}'" if not HAN_RE.search(m.group(1)) else m.group(0), text)
    for full, half in {**CJK_TO_ASCII, **FULLWIDTH_ASCII_TO_ASCII}.items():
        pattern = fr'(?<=[A-Za-z0-9_]){re.escape(full)}(?=[A-Za-z0-9_])'
        text = re.sub(pattern, lambda _m, repl=half: repl, text)

    def normalize_span(match: re.Match[str]) -> str:
        value = match.group(0)
        if (FORMULA_OPERATOR_RE.search(value)
                or any(ch in value for ch in FULLWIDTH_ASCII_TO_ASCII)
                or '（' in value or '）' in value):
            return _normalize_halfwidth(value)
        return value

    return EMBEDDED_TECH_SPAN_RE.sub(normalize_span, text)

def _normalize_chinese(text: str) -> str:
    text = re.sub(fr'["＂]([^"＂\n]*[{HAN}][^"＂\n]*)["＂]', r'“\1”', text)
    text = re.sub(fr"['＇]([^'＇\n]*[{HAN}][^'＇\n]*)['＇]", r'‘\1’', text)
    text = re.sub(fr'\(([^()\n]*[{HAN}][^()\n]*)\)', r'（\1）', text)
    text = re.sub(fr'(?<=[{HAN}])\.{{3,}}(?=$|\s|[{HAN}，。；：！？）》】”])', '……', text)
    text = re.sub(fr'(?<=[{HAN}]),|,(?=[{HAN}])', '，', text)
    text = re.sub(fr'(?<=[{HAN}]);|;(?=[{HAN}])', '；', text)
    text = re.sub(fr'(?<=[{HAN}]):(?=\s|[{HAN}“‘（【《])|(?<=[{HAN}]):$', '：', text)
    text = re.sub(fr'(?<=[{HAN}）】》”’])\?', '？', text)
    text = re.sub(fr'(?<=[{HAN}）】》”’])!', '！', text)
    text = re.sub(fr'(?<=[{HAN}])/(?=[{HAN}])', '／', text)
    text = re.sub(fr'(?<=[{HAN}）】》”’])\.(?=$|\s)', '。', text)
    return _normalize_embedded_nonchinese(text)


def _normalize_halfwidth(text: str) -> str:
    for full, half in FULLWIDTH_ASCII_TO_ASCII.items():
        text = text.replace(full, half)
    for full, half in CJK_TO_ASCII.items():
        text = text.replace(full, half)
    return text


def _normalize_mixed(text: str) -> str:
    # Chinese payload inside ASCII pairs -> full-width Chinese pairs.
    text = re.sub(fr'["＂]([^"＂\n]*[{HAN}][^"＂\n]*)["＂]', r'“\1”', text)
    text = re.sub(fr"['＇]([^'＇\n]*[{HAN}][^'＇\n]*)['＇]", r'‘\1’', text)
    text = re.sub(fr'\(([^()\n]*[{HAN}][^()\n]*)\)', r'（\1）', text)

    # Pure English/version/formula payload inside Chinese pairs -> half-width pairs.
    text = re.sub(r'（([^（）\n]*[A-Za-z0-9][^（）\n]*)）',
                  lambda m: f"({m.group(1)})" if not HAN_RE.search(m.group(1)) else m.group(0), text)
    text = re.sub(r'“([^“”\n]*[A-Za-z0-9][^“”\n]*)”',
                  lambda m: f'"{m.group(1)}"' if not HAN_RE.search(m.group(1)) else m.group(0), text)
    text = re.sub(r'‘([^‘’\n]*[A-Za-z0-9][^‘’\n]*)’',
                  lambda m: f"'{m.group(1)}'" if not HAN_RE.search(m.group(1)) else m.group(0), text)

    return _normalize_embedded_nonchinese(_normalize_chinese(text))


def normalize_fragment(text: str) -> tuple[str, str]:
    kind = context_kind(text)
    if kind == "chinese":
        return _normalize_chinese(text), kind
    if kind in {"english", "technical"}:
        return _normalize_halfwidth(text), kind
    if kind == "mixed":
        return _normalize_mixed(text), kind
    return text, kind


def audit_fragment(text: str) -> list[PunctuationIssue]:
    if not text:
        return []
    normalized, kind = normalize_fragment(text)
    if normalized == text:
        return []
    if kind == "chinese":
        return [PunctuationIssue(
            "halfwidth-punctuation-in-chinese", "error",
            "中文文章／中文语境应使用全角中文标点。", text, normalized)]
    if kind in {"english", "technical"}:
        return [PunctuationIssue(
            "fullwidth-punctuation-in-english-code-formula", "error",
            "英文、代码或公式语境应使用半角标点。", text, normalized)]
    return [PunctuationIssue(
        "contextual-punctuation-mismatch", "warning",
        "中英文混合内容存在与局部语境不一致的标点，请按局部文字类型修正。", text, normalized)]
