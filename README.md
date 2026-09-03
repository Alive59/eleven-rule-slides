# eleven-rule-slides

A Claude skill for building and auditing research presentation slides against the
eleven rules of slide composition: every slide carries a title, two to four summary
sentences, and supporting evidence; numbers are highlighted; chapter separators keep
titles short; red enclosures link a claim to the part of the figure that proves it.

The skill does two things that a prompt alone does not:

- **`scripts/slide_kit.js`** makes the layout rules structural. Zone geometry, the
  two-font split, minimum sizes, and the page number are fixed by the generator, so
  a deck cannot violate them by accident.
- **`scripts/check_rules.py`** audits a finished `.pptx` — including decks the kit
  did not build — and reports what is mechanically wrong, slide by slide.

## Layout

Canvas is `LAYOUT_WIDE`, 13.333 x 7.5 in.

| Zone | x, y, w, h (in) | Type |
|---|---|---|
| Page number | 11.98, 0.18, 0.95, 0.55 | auto field, right-top, 28pt black |
| Title | 0.6, 0.28, 11.5, 0.75 | 32pt bold, one line |
| Summary | 0.6, 1.15, 12.1, ≤ 1.5 | 20pt bullets, 2–4 of them |
| Evidence | 0.6, 2.75, 12.1, 4.35 | figures, tables, charts |

The evidence zone is 58% of the slide height and is reserved. Cover, chapter
separator, and closing slides carry no page number.

## Usage

Generate:

```js
const { createDeck, addContentSlide, fitImages } = require('./scripts/slide_kit.js');

const deck = createDeck({ fontEn: 'Calibri', fontJa: 'Yu Gothic' });
const { slide, zone } = addContentSlide(deck, {
  title: 'Accuracy by input class',
  summary: ['Trained on **1.3M** samples; IoU reached **0.79**.'],
});
fitImages(slide, zone, [{ path: 'fig.png', caption: 'Held-out split' }]);
deck.pres.writeFile({ fileName: 'talk.pptx' });
```

`**bold**` marks a highlighted numeral. Mixed Japanese and Latin text in one string
is split run by run, with the Latin face written to `<a:latin>` and the Japanese
face to `<a:ea>`.

Audit:

```bash
python scripts/check_rules.py talk.pptx
python scripts/check_rules.py talk.pptx --font-en Calibri --font-ja "Yu Gothic"
```

FAIL must be fixed; WARN needs a reason. Exit status is 1 if any FAIL is found, so
it drops into a pre-commit hook or CI step unchanged.

What the audit checks: page-number presence, position, size and color; title and
body point sizes; the two-font rule; summary line count; parenthetical explanation
inside a summary; shapes off-canvas or crowding the edge; missing evidence; a
results slide with no numeral. Zones are inferred per slide, so decks built
elsewhere are checked correctly. Text inside a group, a diagram, or a flow chart is
artwork and is exempt from the font and size rules.

What it cannot check: whether a sentence states a conclusion, whether a title is
specific, whether a figure is large enough to read. Render the deck and look at it.

Retrofitting a deck you did not generate:

```python
from scripts.page_number import add_page_number
add_page_number(slide, n, slide_width_inches)   # skips covers and separators
```

## Content precedence

When a deck already exists, the content on it is the fixed point. Keep its wording,
position, and size; add new material into the space that remains; if it does not
fit, say so rather than displacing what is there. Enlarging body text into a figure
trades a font violation for an overlap.

## Installing as a Claude skill

Zip the repository contents so that `SKILL.md` sits at the root of the archive, and
rename it to `.skill`:

```bash
zip -r eleven-rule-slides.skill SKILL.md references scripts
```

## Requirements

`pptxgenjs` (npm) to generate, `python-pptx` to audit, LibreOffice and `pdftoppm`
to render for visual QA.
