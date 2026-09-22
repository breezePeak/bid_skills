# General layout rules

## Authority

Use this precedence:

1. user-confirmed requirements;
2. provided template;
3. project style guide;
4. fallback rules in this skill.

A template may intentionally differ from generic preferences. Do not "improve" a valid template away.

## Page geometry

Before changing block sizes, calculate printable width and height from the actual section:

- paper width/height;
- portrait/landscape orientation;
- left/right/top/bottom margins;
- section breaks;
- headers/footers when they constrain the visual area.

A block that fits an abstract A4 page can still overflow the real section.

## Paragraphs and headings

Check:

- heading hierarchy is not accidentally skipped;
- body style is not used as a fake heading when a real heading style is available;
- paragraph alignment follows the active authority;
- first-line indentation is applied deliberately and not simulated with spaces;
- line spacing and paragraph spacing are consistent;
- headings are not stranded at the bottom of a page without their following content;
- accidental empty paragraphs do not create large blank areas or blank pages.

## Figures and images

- preserve aspect ratio unless distortion is explicitly intended;
- fit within printable width and height;
- keep captions visually associated with the figure;
- avoid making a figure unreadably small just to keep it on one page;
- for complex diagrams, prefer re-layout over global shrinking.

## Verification

Structural validity is necessary but not sufficient. Always verify the final rendered output for layout-sensitive changes.

## Technical-bid orientation policy

For technical bids, portrait is mandatory for all sections except the section containing `技术偏离表`. Any other landscape section is a blocking layout defect, even when the page width would make an ordinary table easier to fit. Re-layout the table/content instead of rotating that section.
