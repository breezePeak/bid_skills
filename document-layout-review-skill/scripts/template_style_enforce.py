#!/usr/bin/env python3
"""应用已确认的模板样式，并读取本次正文规则。"""
from __future__ import annotations

import argparse
import copy
import json
import re
import zipfile
from pathlib import Path
from lxml import etree
from template_format_contract import StyleResolver, copy_style_dependencies, repair_paragraph, paragraph_style, expected_style_id
from numbering_core import preserve_heading_numbering
from text_rules import (
    body_paragraph_items,
    TextRulesError, load_rules, load_profile, expected_body_props, effective_run_props,
    rules_digest, ensure_rpr, restore_emphasis, apply_explicit_body_overrides,
)
from body_font_exceptions import (
    body_run_font_snapshot,
    clear_fonts_preserving_body_exception,
    restore_body_font_snapshot,
)

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


def clear_run_font_size_overrides(
    p: etree._Element,
    *,
    clear_heading_emphasis: bool = False,
    preserve_body_black: bool = False,
    text_rules: dict | None = None,
) -> int:
    """逐个清理字体/字号覆盖；正文保留项由本次文字规则决定。"""
    removed = 0
    for rpr in p.xpath(".//w:rPr", namespaces=NS):
        tags = ["rFonts", "sz", "szCs"]
        if clear_heading_emphasis:
            tags += ["b", "bCs", "i", "iCs"]
        for tag in tags:
            if tag == "rFonts" and preserve_body_black:
                removed += int(clear_fonts_preserving_body_exception(rpr, text_rules))
                continue
            node = rpr.find(Q(tag))
            if node is not None:
                rpr.remove(node)
                removed += 1
    return removed


def replace_style_definitions(target_styles: bytes, template_styles: bytes, style_ids: set[str]) -> bytes:
    return copy_style_dependencies(target_styles, template_styles, style_ids)


def paragraph_section_map(body: etree._Element) -> dict[etree._Element, int]:
    section = 1
    out: dict[etree._Element, int] = {}
    for child in body:
        if child.tag == Q("p"):
            out[child] = section
            if child.find("w:pPr/w:sectPr", namespaces=NS) is not None:
                section += 1
        elif child.tag == Q("tbl"):
            for p in child.xpath(".//w:p", namespaces=NS):
                out[p] = section
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
    ap.add_argument("--text-rules", type=Path, help="本次文字规则 JSON")
    a = ap.parse_args()

    try:
        if a.out.resolve() in {a.docx.resolve(), a.template.resolve()}:
            raise TextRulesError("输出路径必须与输入文档、模板不同。")
        rules = load_rules(a.text_rules)
        template_profile = load_profile(a.template)
        expected_body_props(template_profile, rules)
    except TextRulesError as exc:
        result = {"status": "failed", "code": "text-rules-invalid", "error": str(exc)}
        if a.json_out:
            a.json_out.parent.mkdir(parents=True, exist_ok=True)
            a.json_out.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 2
    source_profile = load_profile(a.docx)
    contract = load_contract(a.style_json)
    roles = contract["roles"]
    technical_bid = detect_technical_bid(contract, a.technical_bid)
    canonical_style_ids = {sid for sid in roles.values() if sid}

    with zipfile.ZipFile(a.template) as zt, zipfile.ZipFile(a.docx) as zd:
        tnames = set(zt.namelist())
        dnames = set(zd.namelist())
        target_files = {n: zd.read(n) for n in zd.namelist()}
        template_style_root=etree.fromstring(zt.read('word/styles.xml'))
        available={x.get(Q('styleId')) for x in template_style_root.findall(Q('style'))}
        for name in dnames:
            if not name.startswith('word/') or not name.endswith('.xml') or not Path(name).stem.startswith(('document','header','footer','footnotes','endnotes')):
                continue
            story=etree.fromstring(zd.read(name))
            for ref in story.xpath('.//w:pStyle|.//w:rStyle',namespaces=NS):
                if ref.get(Q('val')) in available:canonical_style_ids.add(ref.get(Q('val')))
        target_files["word/styles.xml"] = replace_style_definitions(
            target_files["word/styles.xml"], zt.read("word/styles.xml"), canonical_style_ids
        )
        # 主题字体是样式解析的一部分，直接沿用模板主题。
        if "word/theme/theme1.xml" in tnames and "word/theme/theme1.xml" in dnames:
            target_files["word/theme/theme1.xml"] = zt.read("word/theme/theme1.xml")
        root = etree.fromstring(target_files["word/document.xml"])
        preserved_heading_numbering = preserve_heading_numbering(
            root, zd.read("word/styles.xml"),
            zd.read("word/numbering.xml") if "word/numbering.xml" in dnames else None,
        )
        body = root.find("w:body", namespaces=NS)
        if body is None:
            raise SystemExit("document.xml 缺少 w:body")
        sections = paragraph_section_map(body)
        body_candidates = {p for _, p in body_paragraph_items(root, template_profile, source_profile)}
        expected_format = StyleResolver(zt.read("word/styles.xml"))
        current_format = StyleResolver(target_files["word/styles.xml"])
        extra_format_changes = []
        counts: dict[str, int] = {}
        override_removed = 0
        pstyle_changed = 0
        for p in root.xpath(".//w:p", namespaces=NS):
            role = role_for_paragraph(p, roles, sections.get(p, 1), technical_bid)
            if role is None and p in body_candidates and roles.get("body"):
                role = "body"
            if role is None:
                continue
            sid = roles.get(role)
            if not sid:
                continue
            old = p.find("w:pPr/w:pStyle", namespaces=NS)
            old_sid = old.get(Q("val")) if old is not None else None
            kept_fonts = [
                (r, body_run_font_snapshot(source_profile, old_sid, r, rules))
                for r in p.xpath(".//w:r[w:t]", namespaces=NS)
            ] if role == "body" else []
            kept_emphasis = [
                (r, effective_run_props(source_profile, old_sid, r.find("w:rPr", NS)))
                for r in p.xpath(".//w:r[w:t]", namespaces=NS)
            ] if role == "body" else []
            if old_sid != sid:
                set_pstyle(p, sid)
                pstyle_changed += 1
            preserve_numpr = p.find("w:pPr/w:numPr", namespaces=NS) is not None
            if role not in {"image_paragraph"}:
                override_removed += clear_paragraph_overrides(p, preserve_numpr=preserve_numpr)
                override_removed += clear_run_font_size_overrides(
                    p,
                    clear_heading_emphasis=role.startswith("heading"),
                    preserve_body_black=role == "body",
                    text_rules=rules,
                )
                for run, fonts in kept_fonts:
                    restore_body_font_snapshot(run, fonts)
                for run, snapshot in kept_emphasis:
                    rpr = ensure_rpr(run)
                    restore_emphasis(rpr, snapshot, rules)
                    apply_explicit_body_overrides(rpr, rules)
                    if len(rpr) == 0:
                        run.remove(rpr)
            if role != "image_paragraph":
                changes = repair_paragraph(p, sid, expected_format, current_format)
                if changes:
                    extra_format_changes.append({"role": role, "changes": changes})
            counts[role] = counts.get(role, 0) + 1
        # Cover mapped paragraphs outside the ordinary-body classifier as well.
        # Unknown semantic styles are left for explicit mapping and blocked by audit.
        stories={'word/document.xml':root}
        for name in dnames:
            if name!='word/document.xml' and name.startswith('word/') and name.endswith('.xml') and Path(name).stem.startswith(('header','footer','footnotes','endnotes')):
                stories[name]=etree.fromstring(target_files[name])
        for name,story in stories.items():
            for pi,p in enumerate(story.iter(Q('p')),1):
                if p.xpath('ancestor::w:tc|ancestor::w:drawing|ancestor::w:pict|ancestor::w:txbxContent',namespaces=NS):continue
                sid=expected_style_id(paragraph_style(p,current_format),expected_format,roles)
                if sid in expected_format.styles:
                    extra_format_changes.extend({'part':name,'paragraph':pi,**change} for change in repair_paragraph(p,sid,expected_format,current_format))
            if name!='word/document.xml':
                target_files[name]=etree.tostring(story,xml_declaration=True,encoding='UTF-8',standalone=True)
        target_files["word/document.xml"] = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    a.out.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(a.out, "w", zipfile.ZIP_DEFLATED) as zo:
        for name, data in target_files.items():
            zo.writestr(name, data)
    report = {
        "status": "ok",
        "text_rules": rules,
        "text_rules_sha256": rules_digest(rules),
        "style_contract": str(a.style_json),
        "technical_bid_mode": technical_bid,
        "roles_applied": counts,
        "paragraph_style_changes": pstyle_changed,
        "inherited_heading_numbering_preserved": preserved_heading_numbering,
        "direct_overrides_removed": override_removed,
        "extra_format_changes": extra_format_changes,
    }
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if a.json_out:
        a.json_out.parent.mkdir(parents=True, exist_ok=True)
        a.json_out.write_text(text + "\n", encoding="utf-8")
    print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
