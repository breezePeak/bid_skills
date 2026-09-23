#!/usr/bin/env python3
"""Template-owned table alignment/margins/header policy shared by repair/audit.

Column widths and merges are intentionally excluded. A width repair cannot
silently invent alignment or overwrite the template's cell margins.
"""
from __future__ import annotations

import argparse
import copy
import json
import os
import tempfile
import zipfile
from pathlib import Path
from lxml import etree as E
from template_format_contract import StyleResolver, FormatContractError, Q, NS, parse, serialize, signature, merge

TABLE_PROPS = ('jc', 'tblInd', 'tblCellMar')
CELL_PROPS = ('vAlign', 'tcMar', 'textDirection')
ROW_PROPS = ('tblHeader', 'cantSplit')


def text(node):
    return ''.join(node.xpath('.//w:t/text()', namespaces=NS)).strip()


def headers(tbl):
    row = tbl.find(Q('tr'))
    return tuple(''.join(text(c).split()) for c in row.findall(Q('tc'))) if row is not None else ()


def select_props(props, names):
    result = E.Element(props.tag if props is not None else Q('properties'))
    if props is not None:
        for name in names:
            n = props.find(Q(name))
            if n is not None:
                result.append(copy.deepcopy(n))
    return result


def set_props(owner, kind, names, expected):
    old = owner.find(Q(kind))
    if old is None:
        old = E.Element(Q(kind)); owner.insert(0, old)
    for name in names:
        for n in list(old.findall(Q(name))):
            old.remove(n)
        new = expected.find(Q(name))
        if new is not None:
            old.append(copy.deepcopy(new))
    try:
        from numbering_policy import ordered_properties
    except ImportError:
        ordered_properties = None
    if ordered_properties:
        ordered_properties(old)
    # Table property children also have schema order (the legacy helper only sorts pPr/rPr).
    orders={
        'tblPr':'tblStyle tblpPr tblOverlap bidiVisual tblStyleRowBandSize tblStyleColBandSize tblW jc tblCellSpacing tblInd tblBorders shd tblLayout tblCellMar tblLook tblCaption tblDescription tblPrChange',
        'trPr':'cnfStyle divId gridBefore gridAfter wBefore wAfter cantSplit trHeight tblHeader tblCellSpacing jc hidden ins del trPrChange',
        'tcPr':'cnfStyle tcW gridSpan hMerge vMerge tcBorders shd noWrap tcMar textDirection tcFitText vAlign hideMark headers cellIns cellDel cellMerge tcPrChange'}
    positions={name:i for i,name in enumerate(orders.get(kind,'').split())}
    if positions:
        children=list(old)
        for child in children:old.remove(child)
        for child in sorted(children,key=lambda n:positions.get(E.QName(n).localname,len(positions))):old.append(child)


def normalized(props, names):
    out = []
    for name in names:
        n = props.find(Q(name)) if props is not None else None
        if name == 'vAlign':
            key = n.get(Q('val'), 'top') if n is not None else 'top'
        elif name in ROW_PROPS:
            key = n is not None and n.get(Q('val'), '1') not in {'0', 'false', 'off'}
        elif name == 'jc':
            key = n.get(Q('val'), 'left') if n is not None else 'left'
        elif name == 'textDirection':
            key = n.get(Q('val'), 'lrTb') if n is not None else 'lrTb'
        else:
            key = signature(n)
        out.append((name, key))
    return out


class TableLayoutContract:
    def __init__(self, template: Path, profile: Path | None = None):
        with zipfile.ZipFile(template) as z:
            self.styles = StyleResolver(z.read('word/styles.xml'))
            self.root = parse(z.read('word/document.xml'))
        self.samples = list(self.root.iter(Q('tbl')))
        self.canonical = None
        if profile:
            from runtime_preflight import validate_profile
            validate_profile(template, profile)
            data = json.loads(Path(profile).read_text(encoding='utf-8-sig'))
            self.canonical = data.get('table', {}).get('style_id')
            if self.canonical is not None:
                self.canonical = str(self.canonical)

    def sid(self, tbl):
        n = tbl.find('w:tblPr/w:tblStyle', NS)
        if n is not None:
            return n.get(Q('val'))
        return next((sid for sid, s in self.styles.styles.items() if s.get(Q('type')) == 'table'
                     and s.get(Q('default'), '0') not in {'0', 'false', 'off'}), None)

    def props(self, sample, kind, row=None, col=None):
        result = E.Element(Q(kind))
        sid = self.canonical or self.sid(sample)
        chain = self.styles.chain(sid)
        for style in chain:
            merge(result, style.find(Q(kind)))
        if row is not None and col is not None:
            rows = sample.findall(Q('tr'))
            cells = rows[row].findall(Q('tc'))
            look = sample.find('w:tblPr/w:tblLook', NS)
            mask = int(look.get(Q('val'), '0'), 16) if look is not None else 0
            def flag(name, bit):
                return (look.get(Q(name), '1' if mask & bit else '0') not in {'0', 'false', 'off'}) if look is not None else False
            conditions = ['wholeTable']
            if flag('firstRow', 0x20) and row == 0:
                conditions.append('firstRow')
            if flag('lastRow', 0x40) and row == len(rows) - 1:
                conditions.append('lastRow')
            if flag('firstColumn', 0x80) and col == 0:
                conditions.append('firstCol')
            if flag('lastColumn', 0x100) and col == len(cells) - 1:
                conditions.append('lastCol')
            # Layout-affecting band rules cannot be guessed by a generic row map.
            for style in chain:
                for conditional in style.findall(Q('tblStylePr')):
                    typ = conditional.get(Q('type'), '')
                    prop = conditional.find(Q(kind))
                    if typ.startswith('band') and prop is not None and any(prop.find(Q(n)) is not None for n in CELL_PROPS + ROW_PROPS):
                        raise FormatContractError('template-table-layout-band-mapping', '模板条带有不同布局属性，需要明确行分组映射。')
            for condition in conditions:
                for style in chain:
                    for conditional in style.findall(Q('tblStylePr')):
                        if conditional.get(Q('type')) == condition:
                            merge(result, conditional.find(Q(kind)))
            owner = cells[col] if kind == 'tcPr' else rows[row]
            merge(result, owner.find(Q(kind)))
        else:
            merge(result, sample.find(Q(kind)))
        return result

    def sample_key(self, sample):
        rows = sample.findall(Q('tr'))
        cell = [[normalized(self.props(sample, 'tcPr', ri, ci), CELL_PROPS)
                 for ci, _ in enumerate(row.findall(Q('tc')))] for ri, row in enumerate(rows)]
        header = [normalized(self.props(sample, 'trPr', ri, 0), ROW_PROPS) for ri, row in enumerate(rows) if row.findall(Q('tc'))]
        # Any row/column count may use the same uniform header/body contract.
        def unique_groups(items):
            return [sorted({json.dumps(v, ensure_ascii=False, sort_keys=True) for row in group for v in row})
                    for group in (items[:1], items[1:])]
        return [normalized(self.props(sample, 'tblPr'), TABLE_PROPS), unique_groups(cell),
                header[:1], sorted({json.dumps(v, sort_keys=True) for v in header[1:]})]

    def choose(self, target):
        eligible = [s for s in self.samples if not self.canonical or self.sid(s) == self.canonical]
        matching = [s for s in eligible if headers(target) and headers(target) == headers(s)]
        choices = matching or eligible
        if not choices:
            raise FormatContractError('template-table-layout-missing', '模板缺少可确定的表格布局标准，不能使用待修文档的格式。')
        keys = {json.dumps(self.sample_key(s), sort_keys=True, ensure_ascii=False) for s in choices}
        if len(keys) != 1:
            raise FormatContractError('template-table-layout-ambiguous', '模板有多套表格布局，须明确样表或 table.style_id。')
        return choices[0]

    def cell(self, sample, target, ri, ci):
        rows = sample.findall(Q('tr'))
        indexes = [0] if ri == 0 else list(range(1, len(rows))) or [0]
        exact = headers(sample) == headers(target) and bool(headers(target))
        values = []
        for r in indexes:
            cells = rows[r].findall(Q('tc'))
            indexes_c = [ci] if exact and ci < len(cells) and len(cells) == len(target.findall(Q('tr'))[ri].findall(Q('tc'))) else range(len(cells))
            for c in indexes_c:
                values.append(self.props(sample, 'tcPr', r, c))
        if not values:
            raise FormatContractError('template-table-layout-empty', '模板样表没有可映射单元格。')
        keys = {json.dumps(normalized(v, CELL_PROPS), sort_keys=True) for v in values}
        if len(keys) != 1:
            raise FormatContractError('template-cell-layout-ambiguous', '模板不同单元格布局不同，无法确定当前单元格的对应关系。')
        return values[0]

    def row(self, sample, ri):
        rows = sample.findall(Q('tr'))
        indexes = [0] if ri == 0 else list(range(1, len(rows))) or [0]
        values = [self.props(sample, 'trPr', r, 0) for r in indexes if rows[r].findall(Q('tc'))]
        keys = {json.dumps(normalized(v, ROW_PROPS), sort_keys=True) for v in values}
        if len(keys) != 1:
            raise FormatContractError('template-row-layout-ambiguous', '模板数据行的跨页或重复表头规则不一致，需要明确行映射。')
        return values[0]


def examine(root, contract, repair=False):
    issues = []
    for ti, tbl in enumerate(root.iter(Q('tbl')), 1):
        sample = contract.choose(tbl)
        jobs = [(tbl, 'tblPr', TABLE_PROPS, contract.props(sample, 'tblPr'), {'table': ti})]
        for ri, row in enumerate(tbl.findall(Q('tr'))):
            jobs.append((row, 'trPr', ROW_PROPS, contract.row(sample, ri), {'table': ti, 'row': ri + 1}))
            for ci, cell in enumerate(row.findall(Q('tc'))):
                jobs.append((cell, 'tcPr', CELL_PROPS, contract.cell(sample, tbl, ri, ci), {'table': ti, 'row': ri + 1, 'cell': ci + 1}))
        for owner, kind, names, expected, locator in jobs:
            actual = owner.find(Q(kind))
            if normalized(actual, names) != normalized(expected, names):
                issues.append({'severity': 'error', 'code': 'template-table-layout-mismatch', **locator,
                               'properties': list(names), 'message': '表格布局属性不符合模板。'})
                if repair:
                    set_props(owner, kind, names, expected)
    return issues


def audit(template, target, profile=None):
    contract = TableLayoutContract(Path(template), Path(profile) if profile else None)
    issues = []
    with zipfile.ZipFile(target) as z:
        for part in z.namelist():
            if part.startswith('word/') and part.endswith('.xml') and Path(part).stem.startswith(('document', 'header', 'footer', 'footnotes', 'endnotes')):
                root = parse(z.read(part))
                issues.extend({'part': part, **i} for i in examine(root, contract))
    return {'status': 'failed' if issues else 'passed', 'issues': issues}


def repair_xml(xml, template, profile=None):
    root = parse(xml)
    changes = examine(root, TableLayoutContract(Path(template), Path(profile) if profile else None), repair=True)
    return serialize(root), {'status': 'repaired', 'change_count': len(changes), 'changes': changes}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('template', type=Path); ap.add_argument('target', type=Path)
    ap.add_argument('--template-style-json', type=Path); ap.add_argument('--json-out', type=Path)
    a = ap.parse_args()
    try:
        result = audit(a.template, a.target, a.template_style_json)
    except (FormatContractError, OSError, ValueError, KeyError, zipfile.BadZipFile) as exc:
        result = {'status': 'failed', 'issues': [exc.as_issue() if isinstance(exc, FormatContractError) else
                  {'severity': 'error', 'code': 'template-layout-error', 'message': str(exc)}]}
    payload = json.dumps(result, ensure_ascii=False, indent=2)
    if a.json_out:
        a.json_out.parent.mkdir(parents=True, exist_ok=True); a.json_out.write_text(payload + '\n', encoding='utf-8')
    print(payload)
    return 0 if result['status'] == 'passed' else 2


if __name__ == '__main__':
    raise SystemExit(main())
