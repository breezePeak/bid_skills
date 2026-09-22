# Formatting integrity rules

## Purpose

A template is not satisfied merely because `styles.xml` still contains the original styles. The generated document must actually use those styles without bulk direct-format overrides.

## Required checks

For ordinary body paragraphs:

- use the template's mapped body paragraph style;
- do not apply bold to most or all of a paragraph unless that emphasis is intentional content;
- do not apply italic to most or all of a paragraph unless intentional;
- do not override the template body font or size across most of a paragraph;
- preserve short local emphasis when it affects only a small text span.

For tables:

- use the active template's table style when one exists;
- use the active template's table-text paragraph style when one exists;
- preserve template borders, fills, fonts and theme;
- repair structural layout (widths, margins, alignment, repeat header) without replacing the table's visual language.

## Audit and repair

Audit:

```bash
python3 scripts/template_usage_audit.py <template.docx> <output.docx>
```

Conservative repair of high-confidence bulk body formatting drift:

```bash
python3 scripts/template_usage_repair.py <template.docx> <input.docx> --out <repaired.docx>
```

Do not clear all direct formatting globally. A short bold phrase, inline formula, code span, or deliberate emphasis may be legitimate.
