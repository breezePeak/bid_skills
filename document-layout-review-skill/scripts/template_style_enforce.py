#!/usr/bin/env python3
"""按照已确认的模板样式，把 DOCX 正文实际使用的样式拉回模板基线。

设计目标：
- 模板存在时，不只“检查 styles.xml”，而是让正文、标题、表格、题注真正使用模板样式；
- 尽量保留局部加粗/斜体等有意强调，只清理会遮蔽模板字体与字号的直接格式；
- 默认技术标可直接使用固定 JSON，不重复解析默认模板；
- 技术标模式优先从固定样式 JSON 自动识别，避免调用方漏传参数导致正文未被处理。
"""
from __future__ import annotations

import argparse
import copy
import json
import re
import zipfile
from pathlib import Path
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
Q = lambda n: f"{{{W}}}{n}"


def load_contract(path: Path) -> dict:
    data = json.loads(path.read_text(encoding="utf-8"))
    # 固定模板 JSON
    if "semantic_styles" in data:
        roles = {}
        for role, item in data["semantic_styles"].items():
            if isinstance(item, dict) and item.get("source_style_id"):
                roles[role] = str(item["source_style_id"])
        if data.get("table", {}).get("canonical_text_style"):
            roles["table_text"] = str(data["table"]["canonical_text_style"])
        if data.get("table", {}).get("style_id"):
            roles["table"] = str(data["table"]["style_id"])
        return {"kind": "fixed", "roles": roles, "raw": data}
    # 通用模板 profile
    if "semantic_roles" in data:
        roles = {}
        for role, item in data["semantic_roles"].items():
            if not isinstance(item, dict):
                continue
            sid = item.get("style_id")
            if sid:
                roles[role] = str(sid)
            elif role == "caption" and len(item.get("style_ids", [])) == 1:
                roles[role] = str(item["style_ids"][0])
        return {"kind": "generic", "roles": roles, "raw": data}
    raise ValueError("不支持的模板样式 JSON 格式")


def text_of(p: etree._Element) -> str:
    return "".join(p.xpath(".//w:t/text()", namespaces=NS)).strip()


def ensure_ppr(p: etree._Element) -> etree._Element:
    ppr = p.find(Q("pPr"))
    if ppr is None:
        ppr = etree.Element(Q("pPr"))
        p.insert(0, ppr)
    return ppr


def set_pstyle(p: etree._Element, sid: str) -> None:
    ppr = ensure_ppr(p)
    node = ppr.find(Q("pStyle"))
    if node is None:
        node = etree.Element(Q("pStyle"))
        ppr.insert(0, node)
    node.set(Q("val"), sid)


def clear_paragraph_overrides(p: etree._Element, *, preserve_numpr: bool) -> int:
    ppr = ensure_ppr(p)
    removed = 0
    # 让模板样式控制缩进、对齐、行距和段间距。
    for tag in ("spacing", "jc", "ind", "keepNext", "keepLines", "pageBreakBefore"):
        if preserve_numpr and tag == "ind":
            continue
        node = ppr.find(Q(tag))
        if node is not None:
            ppr.remove(node)
            removed += 1
    return removed


def clear_run_font_size_overrides(p: etree._Element, *, clear_heading_emphasis: bool = False) -> int:
    """清理会覆盖语义样式的 run 字体/字号直接格式。

    注意这里必须逐个 rPr 清理，而不是按段落覆盖率判断。只要某个正文 run
    带有错误字体或字号，就可能在页面上形成局部突变。
    """
    removed = 0
    for rpr in p.xpath(".//w:rPr", namespaces=NS):
        tags = ["rFonts", "sz", "szCs"]
        if clear_heading_emphasis:
            tags += ["b", "bCs", "i", "iCs"]
        for tag in tags:
            node = rpr.find(Q(tag))
            if node is not None:
                rpr.remove(node)
                removed += 1
    return removed


def replace_style_definitions(target_styles: bytes, template_styles: bytes, style_ids: set[str]) -> bytes:
    troot = etree.fromstring(target_styles)
    sroot = etree.fromstring(template_styles)
    # docDefaults 也属于模板字体基线。
    tdd = troot.find(Q("docDefaults"))
    sdd = sroot.find(Q("docDefaults"))
    if sdd is not None:
        if tdd is not None:
            troot.replace(tdd, copy.deepcopy(sdd))
        else:
            troot.insert(0, copy.deepcopy(sdd))
    source = {s.get(Q("styleId")): s for s in sroot.findall(Q("style"))}
    target = {s.get(Q("styleId")): s for s in troot.findall(Q("style"))}
    for sid in style_ids:
        if sid not in source:
            continue
        replacement = copy.deepcopy(source[sid])
        if sid in target:
            old = target[sid]
            troot.replace(old, replacement)
        else:
            troot.append(replacement)
    return etree.tostring(troot, xml_declaration=True, encoding="UTF-8", standalone=True)


def paragraph_section_map(body: etree._Element) -> dict[int, int]:
    section = 1
    out: dict[int, int] = {}
    for child in body:
        if child.tag == Q("p"):
            out[id(child)] = section
            if child.find("w:pPr/w:sectPr", namespaces=NS) is not None:
                section += 1
        elif child.tag == Q("tbl"):
            for p in child.xpath(".//w:p", namespaces=NS):
                out[id(p)] = section
    return out


def role_for_paragraph(
    p: etree._Element,
    roles: dict[str, str],
    section_no: int,
    technical_bid: bool,
) -> str | None:
    # 表格优先。
    if p.xpath("ancestor::w:tc", namespaces=NS):
        return "table_text" if roles.get("table_text") else None
    pstyle = p.find("w:pPr/w:pStyle", namespaces=NS)
    sid = pstyle.get(Q("val")) if pstyle is not None else None
    for name in [f"heading{i}" for i in range(1, 10)]:
        if sid and roles.get(name) == sid:
            return name
    if sid and roles.get("caption") == sid:
        return "caption"
    if sid and roles.get("image_paragraph") == sid:
        return "image_paragraph"
    if sid and roles.get("body") == sid:
        return "body"
    txt = text_of(p)
    if re.match(r"^[图表]\s*\d+(?:[.．、:\s]|$)", txt):
        return "caption" if roles.get("caption") else None
    # 默认技术标中，第 3 节开始才是正式正文；封面和目录不强制改成 Normal。
    if technical_bid and section_no >= 3 and txt and not p.xpath(
        ".//w:drawing|.//w:object|.//w:pict", namespaces=NS
    ):
        return "body" if roles.get("body") else None
    return None


def detect_technical_bid(contract: dict, cli_flag: bool) -> bool:
    """固定模板应由样式 JSON 自动声明文档类型，不依赖调用方记得传参数。"""
    if cli_flag:
        return True
    raw = contract.get("raw") or {}
    template = raw.get("template") if isinstance(raw, dict) else None
    if isinstance(template, dict) and str(template.get("document_type", "")).lower() == "technical_bid":
        return True
    return False


def main() -> int:
    ap = argparse.ArgumentParser(description="按模板样式 JSON 强制正文实际使用模板样式")
    ap.add_argument("template", type=Path)
    ap.add_argument("style_json", type=Path)
    ap.add_argument("docx", type=Path)
    ap.add_argument("--out", type=Path, required=True)
    ap.add_argument("--json-out", type=Path)
    ap.add_argument("--technical-bid", action="store_true")
    a = ap.parse_args()

    contract = load_contract(a.style_json)
    roles = contract["roles"]
    technical_bid = detect_technical_bid(contract, a.technical_bid)
    canonical_style_ids = {sid for sid in roles.values() if sid}

    with zipfile.ZipFile(a.template) as zt, zipfile.ZipFile(a.docx) as zd:
        tnames = set(zt.namelist())
        dnames = set(zd.namelist())
        target_files = {n: zd.read(n) for n in zd.namelist()}
        target_files["word/styles.xml"] = replace_style_definitions(
            target_files["word/styles.xml"], zt.read("word/styles.xml"), canonical_style_ids
        )
        # 主题字体是样式解析的一部分，直接沿用模板主题。
        if "word/theme/theme1.xml" in tnames and "word/theme/theme1.xml" in dnames:
            target_files["word/theme/theme1.xml"] = zt.read("word/theme/theme1.xml")
        root = etree.fromstring(target_files["word/document.xml"])
        body = root.find("w:body", namespaces=NS)
        if body is None:
            raise SystemExit("document.xml 缺少 w:body")
        sections = paragraph_section_map(body)
        counts: dict[str, int] = {}
        override_removed = 0
        pstyle_changed = 0
        for p in root.xpath(".//w:p", namespaces=NS):
            role = role_for_paragraph(p, roles, sections.get(id(p), 1), technical_bid)
            if role is None:
                continue
            sid = roles.get(role)
            if not sid:
                continue
            old = p.find("w:pPr/w:pStyle", namespaces=NS)
            old_sid = old.get(Q("val")) if old is not None else None
            if old_sid != sid:
                set_pstyle(p, sid)
                pstyle_changed += 1
            preserve_numpr = p.find("w:pPr/w:numPr", namespaces=NS) is not None
            if role not in {"image_paragraph"}:
                override_removed += clear_paragraph_overrides(p, preserve_numpr=preserve_numpr)
                override_removed += clear_run_font_size_overrides(
                    p, clear_heading_emphasis=role.startswith("heading")
                )
            counts[role] = counts.get(role, 0) + 1
        target_files["word/document.xml"] = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    a.out.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(a.out, "w", zipfile.ZIP_DEFLATED) as zo:
        for name, data in target_files.items():
            zo.writestr(name, data)
    report = {
        "status": "ok",
        "style_contract": str(a.style_json),
        "technical_bid_mode": technical_bid,
        "roles_applied": counts,
        "paragraph_style_changes": pstyle_changed,
        "direct_overrides_removed": override_removed,
    }
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if a.json_out:
        a.json_out.parent.mkdir(parents=True, exist_ok=True)
        a.json_out.write_text(text + "\n", encoding="utf-8")
    print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
