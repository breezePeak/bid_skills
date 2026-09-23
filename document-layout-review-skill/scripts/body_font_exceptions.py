"""Compatibility helpers for the shared, configurable body-font policy."""
from __future__ import annotations

from text_rules import load_rules, preserved_fonts, effective_run_props

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
FONT_SLOTS = ("ascii", "hAnsi", "eastAsia", "cs")


def qn(name: str) -> str:
    return f"{{{W}}}{name}"


def preserved_body_fonts(fonts: dict[str, str], rules: dict | None = None) -> dict[str, str]:
    return preserved_fonts(fonts, load_rules() if rules is None else rules)


def _merge_fonts(base: dict[str, str], layer: dict[str, str]) -> None:
    for key in FONT_SLOTS:
        if key in layer or key + "Theme" in layer:
            base.pop(key, None)
            base.pop(key + "Theme", None)
    base.update(layer)


def _style_fonts(profile: dict, sid: str | None) -> dict[str, str]:
    chain = []
    seen = set()
    while sid and sid not in seen:
        seen.add(sid)
        style = profile.get("styles", {}).get(sid)
        if not isinstance(style, dict):
            break
        chain.append(style)
        sid = style.get("based_on")
    fonts: dict[str, str] = {}
    for style in reversed(chain):
        _merge_fonts(fonts, (style.get("run") or {}).get("fonts") or {})
    return fonts


def body_run_font_snapshot(profile: dict, paragraph_sid: str | None, run,
                           rules: dict | None = None) -> dict[str, str]:
    rules = load_rules() if rules is None else rules
    props = effective_run_props(profile, paragraph_sid, run.find("w:rPr", NS))
    return preserved_fonts(props.get("fonts") or {}, rules)


def clear_fonts_preserving_body_exception(rpr, rules: dict | None = None) -> bool:
    """Clear the old font override, retaining only the exempt font slots."""
    node = rpr.find("w:rFonts", NS)
    if node is None:
        return False
    fonts = {key: node.get(qn(key)) for key in (*FONT_SLOTS, *(k + "Theme" for k in FONT_SLOTS)) if node.get(qn(key)) is not None}
    keep = preserved_body_fonts(fonts, rules)
    if not keep:
        rpr.remove(node)
        return True
    before = dict(node.attrib)
    for attr in list(node.attrib):
        if attr not in {qn(key) for key in keep} and attr != qn("hint"):
            del node.attrib[attr]
    return before != node.attrib


def restore_body_font_snapshot(run, fonts: dict[str, str]) -> None:
    """Restore only font names; size, bold, italic and other properties are untouched."""
    if not fonts:
        return
    rpr = run.find("w:rPr", NS)
    if rpr is None:
        rpr = run.makeelement(qn("rPr"), {})
        run.insert(0, rpr)
    node = rpr.find("w:rFonts", NS)
    if node is None:
        node = rpr.makeelement(qn("rFonts"), {})
        rpr.insert(1 if len(rpr) and rpr[0].tag == qn("rStyle") else 0, node)
    for key, value in fonts.items():
        node.set(qn(key), value)
        node.attrib.pop(qn(key + "Theme"), None)
