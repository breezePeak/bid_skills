"""Shared, per-task body-text rules for inspection, repair and final audit."""
from __future__ import annotations

import copy
import hashlib
import json
import math
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from template_style_profile import extract_profile

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W, "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
FONT_SLOTS = ("ascii", "hAnsi", "eastAsia", "cs")
FONT_KEYS = (*FONT_SLOTS, *(key + "Theme" for key in FONT_SLOTS))
EMPHASIS = {"bold": ("b", "bCs"), "italic": ("i", "iCs")}
DEFAULT_RULES = Path(__file__).resolve().parent.parent / "assets" / "default-text-rules.json"


class TextRulesError(ValueError):
    """Invalid, conflicting or not safely executable text rules."""


def qn(name: str) -> str:
    return f"{{{W}}}{name}"


def font_key(value: str) -> str:
    # Font-name aliases are identity mappings, not hard-coded preservation rules.
    value = value.strip().lstrip("@").casefold()
    return {"黑体": "simhei", "宋体": "simsun"}.get(value, value)


def _read_json(path: Path) -> dict:
    def unique_keys(pairs):
        out = {}
        for key, value in pairs:
            if key in out:
                raise TextRulesError(f"文字规则存在重复字段：{key}")
            out[key] = value
        return out
    try:
        return json.loads(Path(path).read_text(encoding="utf-8-sig"), object_pairs_hook=unique_keys)
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise TextRulesError(f"无法读取文字规则 {path}：{exc}") from exc


def _validate(data: dict) -> None:
    if not isinstance(data, dict) or set(data) - {"body"}:
        raise TextRulesError("文字规则顶层仅支持 body；未知字段不得忽略。")
    body = data.get("body", {})
    keys = {"preserve_fonts", "forbidden_fonts", "fonts", "size_pt", "bold", "italic"}
    if not isinstance(body, dict) or set(body) - keys:
        raise TextRulesError(f"body 只支持：{', '.join(sorted(keys))}")
    for key in ("preserve_fonts", "forbidden_fonts"):
        if key in body and (not isinstance(body[key], list) or any(
            not isinstance(v, str) or not v.strip().lstrip("@").strip() for v in body[key]
        )):
            raise TextRulesError(f"body.{key} 必须是非空字体名称组成的数组；空数组表示清空。")
    fonts = body.get("fonts", {})
    if not isinstance(fonts, dict) or set(fonts) - set(FONT_SLOTS) or any(
        not isinstance(v, str) or not v.strip().lstrip("@").strip() for v in fonts.values()
    ):
        raise TextRulesError("body.fonts 仅支持 ascii/hAnsi/eastAsia/cs 对应的非空字体名。")
    size = body.get("size_pt")
    if size is not None and (type(size) not in (int, float) or not math.isfinite(size)
                             or size <= 0 or not float(size * 2).is_integer()):
        raise TextRulesError("body.size_pt 必须是正数，最小步长为 0.5 磅；null 表示沿用模板。")
    for key in EMPHASIS:
        if key in body and body[key] not in ("template", "preserve", "forbid", "require"):
            raise TextRulesError(f"body.{key} 仅支持 template/preserve/forbid/require。")


def load_rules(source: Path | str | dict | None = None) -> dict:
    """Merge field-level overrides; omission inherits, [] clears, unknowns fail."""
    defaults = _read_json(DEFAULT_RULES)
    _validate(defaults)
    supplied = _read_json(Path(source)) if isinstance(source, (str, Path)) else copy.deepcopy({} if source is None else source)
    _validate(supplied)
    body = copy.deepcopy(defaults["body"])
    update = supplied.get("body", {})
    for key, value in update.items():
        body[key] = ({**body[key], **value} if key == "fonts" else copy.deepcopy(value))
    banned = {font_key(f) for f in body["forbidden_fonts"]}
    if "preserve_fonts" in update and any(font_key(f) in banned for f in update["preserve_fonts"]):
        raise TextRulesError("同一字体不能同时明确要求保留和禁止。")
    # A user ban cancels a default exemption without requiring redundant fields.
    body["preserve_fonts"] = [f for f in body["preserve_fonts"] if font_key(f) not in banned]
    if any(font_key(f) in banned for f in body["fonts"].values()):
        raise TextRulesError("正文指定字体与禁止字体冲突，请先确认替代字体。")
    return {"body": body}


def rules_digest(rules: dict) -> str:
    data = json.dumps(rules, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return "sha256:" + hashlib.sha256(data).hexdigest()


def write_rules(path: Path, rules: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(rules, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def read_rpr(rpr) -> dict:
    out = {}
    if rpr is None:
        return out
    fonts = rpr.find("w:rFonts", NS)
    if fonts is not None:
        out["fonts"] = {k: fonts.get(qn(k)) for k in FONT_KEYS if fonts.get(qn(k)) is not None}
    for tag, key in (("sz", "size_half_points"), ("szCs", "size_cs_half_points")):
        node = rpr.find(f"w:{tag}", NS)
        if node is not None and node.get(qn("val")) is not None:
            out[key] = node.get(qn("val"))
    for key, tags in EMPHASIS.items():
        for tag, prop in zip(tags, (key, key + "_cs")):
            node = rpr.find(f"w:{tag}", NS)
            if node is not None:
                out[prop] = node.get(qn("val"), "1").casefold() not in {"0", "false", "off"}
    return out


def merge_props(base: dict, layer: dict, *, style_toggle: bool = False) -> None:
    fonts = dict(base.get("fonts") or {})
    for slot in FONT_SLOTS:
        if slot in layer.get("fonts", {}) or slot + "Theme" in layer.get("fonts", {}):
            fonts.pop(slot, None)
            fonts.pop(slot + "Theme", None)
    fonts.update(layer.get("fonts") or {})
    for key, value in layer.items():
        if key == "fonts":
            continue
        if style_toggle and key in {"bold", "italic", "bold_cs", "italic_cs"}:
            if value:
                base[key] = not base.get(key, False)
        else:
            base[key] = value
    if fonts:
        base["fonts"] = fonts


def style_layers(profile: dict, sid: str | None) -> list[dict]:
    chain, seen = [], set()
    while sid and sid not in seen:
        seen.add(sid)
        item = profile.get("styles", {}).get(sid)
        if not isinstance(item, dict):
            break
        chain.append(item)
        sid = item.get("based_on")
    return list(reversed(chain))


def style_props(profile: dict, sid: str | None, *, defaults: bool = True) -> dict:
    out = copy.deepcopy(profile.get("doc_defaults", {}).get("run") or {}) if defaults else {}
    for item in style_layers(profile, sid):
        merge_props(out, item.get("run") or {}, style_toggle=True)
    return out


def load_profile(path: Path) -> dict:
    """Add complex-script emphasis and theme lookup without changing template assets."""
    profile = extract_profile(path)
    with zipfile.ZipFile(path) as z:
        styles = ET.fromstring(z.read("word/styles.xml"))
        for node in styles.findall("w:style", NS):
            sid = node.get(qn("styleId"))
            if sid in profile["styles"]:
                profile["styles"][sid]["run"].update(read_rpr(node.find("w:rPr", NS)))
        dd = styles.find("w:docDefaults/w:rPrDefault/w:rPr", NS)
        profile.setdefault("doc_defaults", {}).setdefault("run", {}).update(read_rpr(dd))
        theme_map = {}
        if "word/theme/theme1.xml" in z.namelist():
            theme = ET.fromstring(z.read("word/theme/theme1.xml"))
            lang = None
            if "word/settings.xml" in z.namelist():
                lang_node = ET.fromstring(z.read("word/settings.xml")).find("w:themeFontLang", NS)
                if lang_node is not None:
                    lang = lang_node.get(qn("eastAsia"))
            script = {"zh-CN": "Hans", "zh-SG": "Hans", "zh-TW": "Hant", "zh-HK": "Hant",
                      "ja-JP": "Jpan", "ko-KR": "Hang"}.get(lang)
            for prefix in ("major", "minor"):
                node = theme.find(f".//a:fontScheme/a:{prefix}Font", NS)
                if node is None:
                    continue
                for suffix, tag in (("Ascii", "latin"), ("HAnsi", "latin"), ("EastAsia", "ea"), ("Bidi", "cs")):
                    font = node.find(f"a:{tag}", NS)
                    value = font.get("typeface") if font is not None else None
                    if not value and tag == "ea" and script:
                        font = node.find(f"a:font[@script='{script}']", NS)
                        value = font.get("typeface") if font is not None else None
                    if value:
                        theme_map[prefix + suffix] = value
        profile["theme_fonts"] = theme_map
    return profile


def resolve_fonts(fonts: dict, profile: dict) -> dict:
    out = dict(fonts)
    for slot in FONT_SLOTS:
        theme = out.get(slot + "Theme")
        if theme:
            value = profile.get("theme_fonts", {}).get(theme)
            if value:
                out[slot] = value
                out.pop(slot + "Theme", None)
            else:
                # Do not mistake a theme's fallback font name for the active font.
                out.pop(slot, None)
    return out


def effective_run_props(profile: dict, paragraph_sid: str | None, rpr) -> dict:
    if not paragraph_sid:
        paragraph_sid = (profile.get("semantic_roles", {}).get("body") or {}).get("style_id")
    out = style_props(profile, paragraph_sid)
    rs = rpr.find("w:rStyle", NS) if rpr is not None else None
    if rs is not None:
        for item in style_layers(profile, rs.get(qn("val"))):
            merge_props(out, item.get("run") or {}, style_toggle=True)
    merge_props(out, read_rpr(rpr))
    out["fonts"] = resolve_fonts(out.get("fonts") or {}, profile)
    return out


def expected_body_props(profile: dict, rules: dict) -> dict:
    sid = (profile.get("semantic_roles", {}).get("body") or {}).get("style_id")
    out = style_props(profile, sid)
    out["fonts"] = resolve_fonts(out.get("fonts") or {}, profile)
    body = rules["body"]
    merge_props(out, {"fonts": body["fonts"]})
    banned = {font_key(f) for f in body["forbidden_fonts"]}
    for slot in FONT_SLOTS:
        value = out.get("fonts", {}).get(slot)
        if value and font_key(value) in banned:
            raise TextRulesError(f"模板正文 {slot} 字体“{value}”被本次规则禁止；请在 body.fonts 中明确替代字体。")
    if body["size_pt"] is not None:
        out["size_half_points"] = out["size_cs_half_points"] = str(int(body["size_pt"] * 2))
    return out


def preserved_fonts(fonts: dict, rules: dict) -> dict:
    body = rules["body"]
    allowed = {font_key(f) for f in body["preserve_fonts"]}
    forbidden = {font_key(f) for f in body["forbidden_fonts"]}
    return {k: v for k, v in fonts.items() if k in FONT_SLOTS and k + "Theme" not in fonts
            and k not in body["fonts"] and font_key(v) in allowed and font_key(v) not in forbidden}


def font_conflicts(actual: dict, expected: dict, rules: dict, used_slots: set[str] | None = None) -> dict:
    """The same decision is used by all audit and repair callers."""
    bad = {}
    exempt = preserved_fonts(actual, rules)
    banned = {font_key(f) for f in rules["body"]["forbidden_fonts"]}
    for slot in FONT_SLOTS:
        theme_key = slot + "Theme"
        if theme_key in actual:
            if banned and (used_slots is None or slot in used_slots):
                bad[theme_key] = {"actual": actual[theme_key], "expected": expected.get(slot),
                                  "reason": "unresolved-theme-font"}
            elif slot in expected or (theme_key in expected and actual[theme_key] != expected[theme_key]):
                bad[theme_key] = {"actual": actual[theme_key], "expected": expected.get(slot, expected.get(theme_key))}
            continue
        value = actual.get(slot)
        if value is None:
            if slot in rules["body"]["fonts"] or (banned and used_slots is not None and slot in used_slots):
                bad[slot] = {"actual": None, "expected": expected.get(slot), "reason": "unresolved-font"}
            continue
        if font_key(value) in banned:
            bad[slot] = {"actual": value, "expected": expected.get(slot), "reason": "forbidden-font"}
        elif slot not in exempt and (
            (slot in expected and font_key(value) != font_key(expected[slot])) or theme_key in expected
        ):
            bad[slot] = {"actual": value, "expected": expected.get(slot, expected.get(theme_key))}
    return bad


def apply_explicit_body_overrides(rpr, rules: dict) -> None:
    """Keep template styles unchanged; task-specific overrides live on body runs."""
    for slot, value in rules["body"]["fonts"].items():
        set_font(rpr, slot, value)
    size = rules["body"]["size_pt"]
    if size is not None:
        for tag in ("sz", "szCs"):
            set_value(rpr, tag, str(int(size * 2)))
    for key, tags in EMPHASIS.items():
        mode = rules["body"][key]
        if mode in {"forbid", "require"}:
            for tag in tags:
                set_value(rpr, tag, "1" if mode == "require" else "0")


def set_value(rpr, tag: str, value: str) -> None:
    node = rpr.find(f"w:{tag}", NS)
    if node is None:
        node = rpr.makeelement(qn(tag), {})
        order = ("rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
                 "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof",
                 "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern",
                 "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd",
                 "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
                 "specVanish", "oMath", "rPrChange")
        rank = {qn(name): idx for idx, name in enumerate(order)}
        index = next((i for i, child in enumerate(rpr)
                      if rank.get(child.tag, len(order)) > rank.get(qn(tag), len(order))), len(rpr))
        rpr.insert(index, node)
    node.set(qn("val"), value)


def set_font(rpr, slot: str, value: str, *, theme: bool = False) -> None:
    node = rpr.find("w:rFonts", NS)
    if node is None:
        node = rpr.makeelement(qn("rFonts"), {})
        rpr.insert(1 if len(rpr) and rpr[0].tag == qn("rStyle") else 0, node)
    node.set(qn(slot + "Theme" if theme else slot), value)
    node.attrib.pop(qn(slot if theme else slot + "Theme"), None)


def restore_emphasis(rpr, snapshot: dict, rules: dict) -> None:
    for key, tags in EMPHASIS.items():
        if rules["body"][key] == "preserve":
            for tag, prop in zip(tags, (key, key + "_cs")):
                set_value(rpr, tag, "1" if snapshot.get(prop, False) else "0")


def ensure_rpr(run):
    rpr = run.find("w:rPr", NS)
    if rpr is None:
        rpr = run.makeelement(qn("rPr"), {})
        run.insert(0, rpr)
    return rpr


def repair_effective_props(profile: dict, paragraph_sid: str | None, run, expected: dict, rules: dict) -> list[str]:
    """Fix inherited conflicts as well as direct ones without editing shared styles."""
    rpr = run.find("w:rPr", NS)
    actual = effective_run_props(profile, paragraph_sid, rpr)
    changed = []
    for key, conflict in font_conflicts(actual.get("fonts") or {}, expected.get("fonts") or {}, rules, run_font_slots(run)).items():
        slot = key.removesuffix("Theme")
        wanted = expected.get("fonts") or {}
        if slot in wanted:
            set_font(ensure_rpr(run), slot, wanted[slot])
        elif slot + "Theme" in wanted and not rules["body"]["forbidden_fonts"]:
            set_font(ensure_rpr(run), slot, wanted[slot + "Theme"], theme=True)
        else:
            raise TextRulesError(f"正文 {slot} 字体无法确定安全替代值；请通过 body.fonts 指定。")
        changed.append("font:" + slot)
    for tag, key in (("sz", "size_half_points"), ("szCs", "size_cs_half_points")):
        wanted = expected.get(key) or (expected.get("size_half_points") if tag == "szCs" else None)
        got = actual.get(key) or (actual.get("size_half_points") if tag == "szCs" else None)
        if wanted is not None and str(got) != str(wanted):
            set_value(ensure_rpr(run), tag, str(wanted))
            changed.append(tag)
    for key, tags in EMPHASIS.items():
        mode = rules["body"][key]
        if mode not in {"forbid", "require"}:
            continue
        wanted = mode == "require"
        for tag, prop in zip(tags, (key, key + "_cs")):
            if bool(actual.get(prop, False)) != wanted:
                set_value(ensure_rpr(run), tag, "1" if wanted else "0")
                changed.append(tag)
    return changed


def run_font_slots(run) -> set[str]:
    """Used only to avoid blocking on an unresolved theme for unused scripts."""
    text = "".join(n.text or "" for n in run.findall(".//w:t", NS))
    slots = set()
    for char in text:
        code = ord(char)
        if char.isspace():
            continue
        if code < 128:
            slots.add("ascii")
        elif 0x2E80 <= code <= 0xD7FF or 0xF900 <= code <= 0xFAFF or 0xFF00 <= code <= 0xFFEF or code >= 0x20000:
            slots.add("eastAsia")
        elif 0x0590 <= code <= 0x08FF or 0xFB1D <= code <= 0xFDFF or 0xFE70 <= code <= 0xFEFF:
            slots.add("cs")
        else:
            slots.add("hAnsi")
    rpr = run.find("w:rPr", NS)
    if rpr is not None and any(rpr.find("w:" + tag, NS) is not None for tag in ("cs", "rtl")):
        slots.add("cs")
    return slots


def body_paragraph_items(root, template_profile: dict, target_profile: dict):
    """Use one body scope for auditing/repairing; exclude tables and named non-body styles.

    In documents with a heading, keep the existing pre-heading cover protection.
    In plain documents without headings, implicit default-style text is still body.
    """
    excluded = set()
    for tag in ("tbl", "txbxContent", "drawing", "pict"):
        for obj in root.findall(f".//w:{tag}", NS):
            excluded.update(obj.findall(".//w:p", NS))
    default_sid = (target_profile.get("semantic_roles", {}).get("body") or {}).get("style_id")
    template_sid = (template_profile.get("semantic_roles", {}).get("body") or {}).get("style_id")

    def classify(p):
        node = p.find("w:pPr/w:pStyle", NS)
        raw = node.get(qn("val")) if node is not None else None
        chain = style_layers(target_profile, raw or default_sid)
        if not chain:
            chain = style_layers(template_profile, raw or template_sid)
        outline = p.find("w:pPr/w:outlineLvl", NS)
        if outline is not None and outline.get(qn("val")) not in {None, "9"}:
            return raw, "heading"
        if any(item.get("paragraph", {}).get("outline_level", 9) < 9 for item in chain):
            return raw, "heading"
        for item in chain:
            name = str(item.get("name", "")).casefold().replace(" ", "")
            if name in {"title", "subtitle", "标题", "副标题", "封面标题"} or name.startswith(("toc", "目录", "caption", "题注", "图题", "表题")):
                return raw, "excluded"
        is_body = (raw or default_sid) in {default_sid, template_sid} or any(
            item.get("id") in {default_sid, template_sid} for item in chain
        )
        return raw, "body" if is_body else "other"

    records = [(i, p, classify(p)) for i, p in enumerate(root.findall(".//w:p", NS), 1) if p not in excluded]
    has_heading = any(kind == "heading" for _, _, (_, kind) in records)
    in_body = False
    for idx, p, (raw_sid, kind) in records:
        if kind == "heading":
            in_body = True
            continue
        if kind == "excluded":
            continue
        if kind == "body" and (raw_sid is not None or not has_heading):
            in_body = True
        if in_body:
            yield idx, p
