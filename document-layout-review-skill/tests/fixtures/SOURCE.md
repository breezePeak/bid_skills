# Fixture provenance

`approved-default-numbering.xml` is the exact decompressed `word/numbering.xml`
from `assets/default-template.docx` in `breezePeak/bid_skills`, commit
`0611bfa277dd71eaf5d5e0b47147b9c3d0b61a17`, blob
`351194d8905fa9b4fb457c10d7f50b292f74fe59`.

The local ZIP-entry CRC was verified during extraction. This file is a regression
fixture, not a replacement for reading the active user's actual template.
Levels 1–9 are decimal, `%1`, `%1.%2`, …, suffix `space`; there is no `第…章` prefix.
Other synthetic test documents use python-docx and small generated bitmap fixtures.
