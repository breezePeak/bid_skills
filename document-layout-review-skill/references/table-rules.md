# Table layout rules

## Core rule

A table is complete only when it is **readable and visually coordinated in the final rendered Word page**.

Passing XML validity, keeping a table style ID, or merely staying inside page margins is not enough. A table is still wrong when one column is unnecessarily narrow, another is unnecessarily wide, headers are broken into awkward fragments, long text is trapped in a compact column, repeated grouping text should have been merged, or Word autofit stretches the page because of one long unbreakable token.

## Template authority

The active template is the formatting standard for every target table; the source
Word supplies content, not an alternative design. Existing target formatting must
be corrected even when it looks intentional. A template with no header fill or
no borders defines an explicit absence, not an unspecified property.

Run the existing template text step, then `template_table_style.py repair` and its
`audit`. Apply the template's table style, fills, borders and text colors. Resolve
its style inheritance/conditional formatting and remove conflicting direct,
paragraph/run, or row-exception fills. The template's explicit canonical table
style applies to ordinary tables even when their headers/column counts differ.
Do not copy a sample's physical grid merely to adopt its appearance. Where an
uploaded template genuinely contains different table styles, use its corresponding
sample/header or declared canonical style; do not choose the most frequent style.

Only after template conformance passes, snapshot the normalized table and perform
width repair. Keep the **normalized template appearance** during that later step.
Do not preserve an original gray header when the template header is unshaded; do
not remove the template's intended color under a blanket “no color” rule.

Template-specific structural contracts remain separate. Existing width proportions
are a reference when the sample has different content, unless the template/user
explicitly fixes those widths. Continue using content-aware reflow and visual QA.

```bash
python scripts/template_table_style.py repair input.docx --template active-template.docx \
  --template-style-json active-template-style.json --out template-tables.docx --json-out changes.json
python scripts/template_table_style.py audit template-tables.docx --template active-template.docx \
  --template-style-json active-template-style.json --json-out template-check.json
```

These steps are built into inspection, repair, final audit and release. The
ordinary `table-layout` guard still prevents later recoloring; it must not be used
to block the preceding authorized template normalization.

## Required inspection order

For every table:

1. identify the header and semantic purpose of every column;
2. determine the current section's printable width;
3. read actual physical grid widths;
4. measure header pressure and body pressure per column;
5. compare actual width allocation with content pressure;
6. detect cramped columns and one-character-per-line risk;
7. detect unjustified abrupt width differences between adjacent columns;
8. detect long unbreakable strings that may force autofit/page stretching;
9. inspect consecutive repeated grouping cells and identify vertical-merge candidates;
10. render the final DOCX and visually verify the result.

Run:

```bash
python3 scripts/table_layout_audit.py <input.docx> \
  --template-style-json <active-template-style.json> \
  --json-out <report.json>
```

The program is a deterministic evidence generator, not an aesthetic oracle. `error` must be fixed. `review` means the Agent must inspect semantics and decide; it must not be ignored or mechanically auto-applied.

## Width model

Use the **current section's printable width** as the hard ceiling:

`paper width - left margin - right margin`

For layout-sensitive DOCX:

- prefer an explicit `dxa` table width;
- prefer fixed table layout after widths are resolved;
- keep `tblGrid` and owning `tcW` values consistent;
- do not rely on Word autofit to rescue poor widths;
- do not allow a single long line to expand the table beyond the printable width.

## Column allocation must be content-aware

Do not use equal widths by default.

Width allocation must consider all of the following:

- header text length;
- semantic role of the column;
- average body content amount;
- longer-body-content pressure such as p90;
- maximum body content pressure;
- actual wrapping behavior after rendering.

Typical roles:

- **compact**: 序号、编号、状态、是否、数量、单位、偏离程度、日期、时间；
- **medium**: 名称、类别、对象、阶段、模块、工作包、责任单位；
- **narrative**: 要求、响应内容、措施、说明、备注、依据、风险、影响、处理方法、输入、输出、结论。

Compact columns should stay compact **but readable**. Medium columns should not be squeezed just because they are not narrative. Narrative columns should receive enough width to avoid dense, awkward wrapping.

### Header is part of width pressure

A short body column can still need more width if the header is long.

Do not create layouts where:

- a four-to-six-character header becomes nearly one character per line;
- a compact-looking header hides very long body content;
- a narrow first column forces every body value into tall stacked text;
- a neighboring low-content column consumes a large share of the page for no reason.

### Visual coordination

The table fails visual balance when:

- one column is conspicuously narrow while a neighboring low-pressure column is very wide;
- width differences are much larger than differences in header/body content pressure;
- short columns contain large unused horizontal space while long-text columns are visibly cramped;
- the table looks like unrelated widths were copied from another document.

The deterministic audit flags severe mismatches, but the Agent must still inspect the rendered page.

## One-character-per-line and cramped-column failures

A descriptive or grouping column that effectively allows only one or two Chinese characters per line is a layout failure when its content is longer than a trivial value.

Do not solve this by globally shrinking the table font.

Preferred repair order:

1. reclaim space from over-wide low-pressure columns;
2. rebalance the entire table;
3. allow natural wrapping;
4. only change font size if the active template explicitly allows it.

## Long unbreakable strings and page stretching

Long URL/path/token/version/hash/identifier strings can cause Word autofit to expand a column or table unexpectedly.

When such strings exist:

- do not let autofit decide the final page width;
- use controlled/fixed table width;
- keep the table inside the printable width;
- where semantically safe, use break opportunities or a display form that does not alter the underlying meaning.

A single line must never be allowed to make the page appear “infinitely wide”.

## Vertical merging of repeated grouping cells

Repeated text is **not automatically** a merge instruction.

First decide what the column means.

Columns that commonly represent a grouping dimension include:

- 标的名称；
- 项目名称；
- 工作对象；
- 类别；
- 阶段；
- 模块；
- 工作包；
- 责任单位/部门；
- 成果类型。

When consecutive rows:

1. contain the same non-empty text in such a grouping column;
2. belong to the same semantic object/group;
3. differ in subordinate/detail columns;

then the repeated grouping cell should normally be vertically merged.

Do **not** automatically merge:

- 序号/编号；
- 状态；
- 是否；
- 偏离程度；
- 数量/单位；
- 日期/时间；
- result/conclusion fields;
- any rows that are separate business records even if the displayed text happens to be equal.

After merge:

- preserve every other cell's content;
- preserve row-to-row correspondence in all other columns;
- center merged text vertically where the template requires it;
- verify borders and page breaks;
- re-render the table.

`table_layout_audit.py` emits these as `review` candidates because semantic ownership cannot be proven solely from XML.

## Alignment

For technical bids:

- every table cell is vertically centered when the template requires it;
- header cells are horizontally centered;
- compact body columns are normally centered;
- narrative body columns are normally left aligned;
- table paragraphs use zero first-line indent;
- table text uses the active template's table-text style.

## Cell margins

Use deliberate internal cell margins. Text must not touch borders. Keep margins consistent across rows.

## Pagination

- repeat the header row on continued pages;
- allow long rows to split when necessary;
- do not create giant blank pages merely to keep a long row together;
- do not rotate ordinary technical-bid tables to landscape;
- only the template-approved section may use landscape orientation.

## Hard failures

A table whose fills/borders/text colors differ from the active template fails even if its original appearance was preserved.

A table fails when any of the following is true:

- total width exceeds printable width;
- the active template's required physical grid structure is violated;
- a meaningful column is so narrow that text is effectively one character per line;
- an autofit table contains a long unbreakable token that can stretch the page;
- long narrative content is trapped in a compact-width column while low-pressure columns are obviously over-wide;
- text touches borders;
- repeated page sections have inconsistent column widths;
- required vertical centering is missing;
- font is aggressively shrunk to compensate for bad layout;
- table orientation is changed just to hide width problems;
- unrequested header/row fills, border colors or text colors are introduced;
- a template style is substituted without an explicit mapping for this table.

## Review-level findings

These require Agent judgment rather than blind automatic repair:

- moderate column-balance mismatch;
- abrupt but potentially justified adjacent width differences;
- repeated grouping cells that may need vertical merging;
- unusual tables whose semantic role cannot be inferred from the header.

For every review-level finding, the Agent must either:

1. repair it; or
2. explicitly determine from content that no repair is appropriate.

Then render the latest DOCX and verify the affected table pages.
