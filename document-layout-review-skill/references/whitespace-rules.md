# Whitespace normalization rules

The goal is to remove **useless** whitespace, not all whitespace.

## Safe automatic candidates

For ordinary Chinese prose and table cells, these are normally safe to remove or normalize:

- ASCII/full-width spaces at paragraph start or end;
- multiple consecutive ordinary spaces that do not express intentional alignment;
- spaces between adjacent Han characters;
- spaces immediately before Chinese punctuation;
- spaces immediately after opening Chinese punctuation or before closing punctuation;
- accidental leading/trailing spaces inside a cell;
- repeated empty paragraphs with no visual purpose.

## Preserve unless clearly accidental

Do not automatically remove spaces in:

- English prose (`OpenAI API`);
- numbers and units (`100 GB`, `5 mm`);
- version/product names;
- URLs and email addresses;
- file-system paths;
- source code and inline code;
- command lines;
- preformatted text;
- identifiers where spaces are meaningful;
- manual form blanks where spacing is part of a visual field.

## CJK punctuation guidance

For ordinary Chinese prose, avoid accidental spaces around punctuation such as:

`，。；：！？、）】》」』`

and after opening punctuation such as:

`（【《「『`

Do not blindly rewrite quoted source text or code-like content.

## DOCX caution

Visible text can be split across multiple Word runs. A phrase that looks contiguous on screen may live in several `<w:r>` elements. Do not rebuild an entire paragraph into one run merely to delete a space because that can destroy mixed formatting, hyperlinks, fields, tracked changes, or comments.

For existing DOCX files:

- automatically repair only cases that can be changed without losing run-level semantics;
- otherwise report the exact paragraph/cell and let the editable source or a document-aware repair routine handle it;
- verify the rendered result after repair.
