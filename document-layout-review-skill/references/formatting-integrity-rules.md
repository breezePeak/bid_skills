# Formatting integrity rules

## Purpose

A template is not satisfied merely because `styles.xml` still contains the original styles. The generated document must actually use those styles without direct-format overrides that change the visible result.

## Required checks

For ordinary body paragraphs:

- use the template's mapped body paragraph style;
- do not apply bold to most or all of a paragraph unless that emphasis is intentional content;
- do not apply italic to most or all of a paragraph unless intentional;
- do not override the template body font or size;
- **font and size are checked per visible run, not by paragraph coverage**: a one-word or one-sentence font/size override that conflicts with the body baseline is still an error;
- character styles that change the ordinary body font/size are also treated as font/size overrides;
- preserve short local emphasis only for semantic emphasis such as bold/italic; “short local span” is not a justification for keeping a conflicting font or size.

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

The audit is a release gate. Any run-level font/size conflict in ordinary body text must fail the gate.

Repair of body formatting drift (run-level font/size conflicts are fixed even when local; bold/italic remains conservative):

```bash
python3 scripts/template_usage_repair.py <template.docx> <input.docx> --out <repaired.docx>
```

Do not clear all direct formatting globally. A short bold phrase, inline formula, code span, or deliberate emphasis may be legitimate. Font/size, however, must still match the paragraph's semantic baseline unless a separately recognized semantic object explicitly requires another style.
