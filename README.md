# eleven-rule-slides

A Claude skill for building and auditing research presentation slides against the
eleven rules of slide composition: every slide carries a title, two to four summary
sentences, and supporting evidence; numbers are highlighted; chapter separators keep
titles short; red enclosures link a claim to the part of the figure that proves it.

The skill does two things that a prompt alone does not:

- **`scripts/slide_kit.py`** (python-pptx) and **`scripts/slide_kit.js`**
  (pptxgenjs) make the layout rules structural. Zone geometry, the two-font split,
  minimum sizes, and the page number are fixed by the generator, so a deck cannot
  violate them by accident. The two kits are ports of each other: same geometry,
  same output under the audit — use whichever the sandbox has.
- **`scripts/check_rules.py`** audits a finished `.pptx` — including decks neither
  kit built — and reports what is mechanically wrong, slide by slide.

It runs under Claude (`SKILL.md`), under Codex and other coding agents
(`AGENTS.md`), and as a ChatGPT Custom GPT (`gpt/instructions.md`).

## Layout

Canvas is `LAYOUT_WIDE`, 13.333 x 7.5 in.

| Zone | x, y, w, h (in) | Type |
|---|---|---|
| Page number | 11.98, 0.18, 0.95, 0.55 | field on the slide master, right-top, 28pt black |
| Title | 0.6, 0.28, 11.5, 0.75 | 32pt bold, one line |
| Summary | 0.6, 1.15, 12.1, ≤ 1.5 | 20pt bullets, 2–4 of them |
| Evidence | 0.6, 2.75, 12.1, 4.35 | figures, tables, charts |

The evidence zone is 58% of the slide height and is reserved. The page number
lives on the slide master, so every slide inherits it; only the title and
acknowledgement slides go without, via a layout that hides master shapes.

## Usage

Generate, in Python:

```python
import sys; sys.path.insert(0, 'scripts')
from slide_kit import create_deck, add_content_slide, fit_images

deck = create_deck(font_en='Calibri', font_ja='Yu Gothic')
slide, zone = add_content_slide(
    deck,
    title='Accuracy by input class',
    summary=['Trained on **1.3M** samples; IoU reached **0.79**.'],
)
fit_images(slide, zone, [{'path': 'fig.png', 'caption': 'Held-out split'}])
deck.save('talk.pptx')
```

or in Node:

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

Titles written as sentences (ending punctuation, a Japanese predicate ending, a
finite verb) and overlong titles are flagged as well.

What it cannot check: whether a sentence states a conclusion, whether a title is
specific, whether a figure is large enough to read. Render the deck and look at it.

Retrofitting a deck you did not generate:

```python
import sys; sys.path.insert(0, 'scripts')
from page_number import add_page_number
add_page_number(slide, n, slide_width_inches)   # skips covers and separators
```

## Content precedence

When a deck already exists, the content on it is the fixed point. Keep its wording,
position, and size; add new material into the space that remains; if it does not
fit, say so rather than displacing what is there. Enlarging body text into a figure
trades a font violation for an overlap.

## Installing

**Claude skill** — zip the repository contents so that `SKILL.md` sits at the root
of the archive, and rename it to `.skill`:

```bash
zip -r eleven-rule-slides.skill SKILL.md references scripts requirements.txt
```

Or clone it into the skills directory:

```bash
git clone https://github.com/Alive59/eleven-rule-slides.git ~/.claude/skills/eleven-rule-slides
```

**Codex and other coding agents** — clone the repository; `AGENTS.md` at the root is
the entry point and points at `SKILL.md`.

**ChatGPT Custom GPT** — `gpt/README.md` has the setup: `gpt/instructions.md` into
the Instructions box, the scripts and references into Knowledge, Code Interpreter
enabled.

## Requirements

`python-pptx` for the audit and the Python generator; `Pillow` for aspect-correct
figure placement; `pptxgenjs` (npm) only for the Node generator; LibreOffice and
`pdftoppm` to render for visual QA.

```bash
pip install -r requirements.txt
npm install pptxgenjs   # only if you use the Node kit
```

ChatGPT Code Interpreter has `python-pptx` and `Pillow` but no Node and no
LibreOffice, which is why the Python kit exists and why a GPT cannot do the visual
QA pass.

## Layout of the repository

```
SKILL.md                    the specification (Claude skill entry point)
AGENTS.md                   entry point for Codex and other coding agents
gpt/instructions.md         condensed build for a ChatGPT Custom GPT
gpt/README.md               how to wire up the Custom GPT
references/eleven-rules.md  the eleven rules, with a do/don't per rule
references/toolchain.md     pptxgenjs and python-pptx footguns, render QA
scripts/slide_kit.py        generator (python-pptx)
scripts/slide_kit.js        generator (pptxgenjs)
scripts/check_rules.py      audit any .pptx against the mechanical rules
scripts/page_number.py      master page number; retrofit onto a deck you did not build
examples/example_deck.py    runnable example (Python)
examples/example_deck.js    runnable example (Node)
tests/test_python_kit.py    smoke tests for the Python kit
```

Run the tests with `pytest tests/` or `python tests/test_python_kit.py`.
