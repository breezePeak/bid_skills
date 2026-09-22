# Table layout rules

## Core rule

A table is complete only when it is **readable in the final rendered Word page**. Preserving a table style ID is not enough. Bad column widths, one-character-per-line wrapping, excessive blank pages, top-aligned cells or stale autofit widths are layout failures.

## Template authority

When an active template exists:

1. keep the template's table style, borders, fills, fonts and table-text style;
2. if the template contains a table with the same header signature, reuse that table's column-width proportions;
3. otherwise derive widths from content pressure while keeping the template visual style;
4. structural layout repairs are allowed and required when the template's old table structure is broken or incompatible with actual content.

Run:

```bash
python3 scripts/table_layout_repair.py <input.docx> --template <template.docx> --output <repaired.docx>
```

The repair uses fixed DOCX widths and writes the same widths to `tblGrid` and every owning `tcW`.

## Width model

Use the **current section's printable width** (paper width minus left/right margins) as the table width ceiling.

For layout-sensitive DOCX:

- set an explicit `dxa` table width;
- set `w:tblLayout w:type="fixed"`;
- set every grid column width;
- set every cell width consistently with its logical column or grid span;
- center the table on the page;
- do not rely on Word autofit to rescue bad widths.

## Column allocation

Do not use equal widths by default.

Classify columns by role and pressure:

- compact: sequence, status, yes/no, date, quantity, deviation degree;
- medium: name, category, object, phase, interface;
- narrative: requirements, response, measures, notes, basis, impact, explanation.

Narrative columns receive more width. Compact columns stay narrow but must remain readable. A descriptive column that collapses into one-character-per-line wrapping is a hard failure.

### Technical deviation table

The technical deviation table is a six-column semantic table:

1. 序号
2. 标的名称
3. 招标技术要求
4. 投标响应内容
5. 偏离程度
6. 备注

The bundled technical-bid template contains the approved six-column structure and width proportions. Reuse those proportions when that template is active. Do not preserve the legacy hidden spacer grid column.

## Alignment

For technical bids:

- every table cell is vertically centered;
- header cells are horizontally centered;
- compact body columns are horizontally centered;
- narrative body columns are left aligned;
- table paragraphs use zero first-line indent;
- table text uses the active template's table-text style.

## Cell margins

Use deliberate internal cell margins. Text must not touch borders. Keep margins consistent across rows.

## Pagination

- repeat the header row on every continued page;
- allow long rows to split when necessary;
- do not create giant blank pages merely to keep a long row together;
- do not rotate ordinary technical-bid tables to landscape;
- only the technical deviation table section may be landscape.

## Hard visual failures

A table fails when any of the following is visible:

- one-character-per-line wrapping caused by a narrow column;
- long narrative content trapped in a compact-width column;
- large numbers of nearly empty continuation pages caused by a bad width allocation;
- text touching borders;
- inconsistent column widths between repeated pages;
- body cells visibly top-aligned when vertical centering is required;
- table width exceeds the section's printable width;
- font is aggressively shrunk to compensate for bad layout;
- table is rotated to landscape outside the technical deviation section.

After structural repair, render the DOCX and inspect the table pages. XML checks alone are not sufficient.
