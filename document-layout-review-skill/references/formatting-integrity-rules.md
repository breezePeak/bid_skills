# Formatting integrity rules

## Required checks

For ordinary body paragraphs:

Read the per-task JSON through `--text-rules`; see `text-rules.md`. Defaults preserve existing `黑体` / `SimHei` fonts, but an explicit user ban or replacement font takes precedence. These rules apply only to ordinary body text, not headings, tables or captions.

- use the template's mapped body paragraph style;
- check bold and italic against the active task rules; default `template` mode rejects conflicting direct emphasis while preserving permitted character-style emphasis;
- apply explicit per-task font/size overrides before falling back to the template; only preserve fonts allowed by the active rules;
- **font and size are checked per visible run, not by paragraph coverage**: any non-exempt font or size conflict is still an error even in a short span;
- character styles that introduce non-exempt font/size conflicts are also treated as font/size overrides;
- preserve emphasis only when permitted by the active rules; “short local span” is not a justification for keeping a conflicting font or size.

For tables:

- use the active template's table style when one exists;
- use the active template's table-text paragraph style when one exists;
- preserve template borders, fills, fonts and theme;
- repair structural layout (widths, margins, alignment, repeat header) without replacing the table's visual language.

## Audit and repair

Audit:

```bash
python3 scripts/template_usage_audit.py <template.docx> <output.docx> --text-rules <text-rules.json>
```

The audit is a release gate. Any non-exempt run-level font or size conflict in ordinary body text must fail the gate.

Repair using the same rules:

```bash
python3 scripts/template_usage_repair.py <template.docx> <input.docx> --out <repaired.docx> --text-rules <text-rules.json>
```

Do not clear all direct formatting globally. A short bold phrase, inline formula, code span, or deliberate emphasis may be legitimate. Font/size and emphasis must match the active rules. Report unsupported or conflicting requirements rather than silently falling back to defaults.
