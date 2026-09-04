---
name: eleven-rule-slides
description: Build research presentation slides that follow the eleven rules of slide composition — title + 2-4 summary sentences + supporting evidence per slide, chapter separator slides, red enclosures linking claims to figures, large fonts (title 32pt and up, body 20pt and up), Calibri for English and Yu Gothic for Japanese, auto page number top-right. Use this skill whenever the user asks for slides, a deck, a presentation, a talk, a progress report, a seminar/lab meeting/conference/defense presentation, or a .pptx file, and also when they ask to review, restructure, or fix existing slides. Applies even if they don't mention the eleven rules by name.
---

# Eleven-rule research slides

Slides built under this skill are read by an audience that does not know the work.
Every slide must answer, on its own: what is this, what did it show, and where can I
see that. The eleven rules exist to force that; `references/eleven-rules.md` has the
full text plus the do/don't pattern for each one. Read it before drafting content.

If your runtime ships a general pptx skill, read it too — in Claude environments
that is `/mnt/skills/public/pptx/SKILL.md`, which carries the render-to-image QA
loop and the library footguns. This skill layers structure and typography on top of
it; it does not replace it. Where no such file exists, `references/toolchain.md` has
the footguns for both generators in short form.

## Runtimes

The skill runs anywhere that has one of the two generators. Pick by what the
sandbox has, not by preference — the two produce the same geometry and both audit
clean under the same `check_rules.py`.

| Runtime | Generator | Notes |
|---|---|---|
| Claude Code / Claude skills | `scripts/slide_kit.js` or `scripts/slide_kit.py` | both toolchains available |
| ChatGPT Code Interpreter, Custom GPT | `scripts/slide_kit.py` | no Node and no network there; `python-pptx` is present |
| Codex and other coding agents | either | `AGENTS.md` at the repo root is the entry point |
| Plain Python pipeline | `scripts/slide_kit.py` | `pip install -r requirements.txt` |

The Python kit cannot define a slide master (python-pptx has no API for one), so it
writes a real `slidenum` field onto each numbered slide instead. PowerPoint still
renumbers the field itself, so reordering slides is safe. Everything else — zone
geometry, the two-font split, minimum sizes, which slides carry no number — is
identical.

## Content precedence

Content already on the slide comes first. When a deck, a draft, or a set of figures
already exists, that material is the fixed point: keep its wording, its position,
and its size, and build everything else around it. Only after it is placed do you
add what the user asked for in this turn, into the space that is left.

This means, in order:

1. Keep what is there. Do not rewrite a sentence, move a figure, restyle a shape, or
   drop a slide that the user did not ask you to touch — even when it violates a
   rule. Report the violation and let them decide.
2. Place the new content the user asked for, fitting it to the room that remains.
3. If the new content does not fit, say so and offer the choice — shorten the new
   material, or change something existing. Never silently displace existing content
   to make room, and never resize an existing figure to fit new text.

The reason is that you cannot see what the existing layout is carrying. A figure at
an odd size may be matched to a companion slide; a sentence you find wordy may be
the wording a reviewer asked for. Enlarging body text into a figure, as happened on
slide 18 of a real deck, trades a font violation for an overlap — worse on the
screen and not your call to make.

The same order applies within one slide: title, then the summary sentences that are
already written, then evidence, then anything new.

## Workflow

0. **Ask which fonts to use** before generating anything, unless the user has
   already said or the deck being edited already has a consistent pair. One short
   question, offering the default: Calibri for English and Yu Gothic for Japanese.
   Pass the answer as `createDeck({ fontEn, fontJa })` — `create_deck(font_en=,
   font_ja=)` in Python — and audit with
   `--font-en` / `--font-ja` so the check matches what was asked for. The rule the
   skill enforces is that a deck is consistent about its two faces, not that the
   faces are any particular pair — a lab or venue template may specify others.
1. **Collect the content before touching code.** For each planned slide write, in
   plain text: the title keyword, the 2-4 summary sentences, and which figure or
   table is the evidence. If a summary sentence has no evidence, or a figure has no
   sentence, the slide is not ready.
2. **Draft the outline with chapter separators** (rule 5), then confirm with the user
   if the deck is longer than ~8 slides or the talk length is unknown.
3. **Generate with `scripts/slide_kit.js` or `scripts/slide_kit.py`.** Both fix the
   geometry, the fonts, and the page number so those rules cannot be violated by
   accident. Use the Python kit when the sandbox has no Node.
4. **Audit with `scripts/check_rules.py deck.pptx`.** Fix every FAIL, then re-run.
   On an existing deck, fix only what content precedence allows — the rest is a
   report, not a task.
5. **Visual QA** — render to images per the pptx skill and look at every slide.
   Overflow and figure placement are the two things the audit cannot see.

The sandbox renderer usually lacks Calibri and Yu Gothic and sometimes drops bullet
glyphs, so judge the preview on layout, overflow, and figure size — not on the exact
glyph shapes. The audit reads the XML, so it is the authority on fonts and sizes.

Dependencies: `python-pptx` always (the audit needs it, and the Python kit is built
on it); `pptxgenjs` (npm) only for the JS kit; `Pillow` for aspect-correct image
placement from the Python kit; LibreOffice + `pdftoppm` for the render. `pip install
-r requirements.txt` covers the Python side.

## Fixed layout (LAYOUT_WIDE, 13.333 x 7.5 in)

| Zone | x, y, w, h (in) | Type |
|---|---|---|
| Page number | 11.98, 0.18, 0.95, 0.55 | auto field, right-top, 28pt pure black |
| Title | 0.6, 0.28, 11.5, 0.75 | 32pt bold, one line |
| Summary | 0.6, 1.15, 12.1, <= 1.5 | 20pt bullets, 2-4 of them |
| Evidence | 0.6, 2.75, 12.1, 4.35 | figures, tables, charts |

The evidence zone is 58% of the slide height and is reserved. Do not let summary
text grow into it — if the sentences do not fit in 1.5", they are too long or too
many, which is rule 6 telling you to split the slide. Equally, do not leave the
zone half empty: a slide with 4.35" of white space below three bullets is a slide
whose figure is too small to read (rule 3).

When retrofitting an existing deck, do not enlarge body text blindly. Estimate the
block's height at the new size first and compare it to the space above the first
figure; where it does not fit, leave the size alone and report that the sentences
need cutting (see Content precedence).

Place figures with `fitImages()`, which spreads them across the zone with `contain`
sizing so nothing is stretched. Aim for each figure to occupy at least 3" of height
when there are one or two of them.

## Typography

- Two faces per deck: one for Latin text, one for Japanese. The default pair is
  **Calibri** and **Yu Gothic**; ask the user first (workflow step 0) and pass
  theirs to `createDeck({ fontEn, fontJa })` / `create_deck(font_en=, font_ja=)`.
  Both kits split mixed strings run-by-run automatically — pass one string and it handles the switch, writing the
  Latin face to `<a:latin>` and the Japanese face to `<a:ea>` so PowerPoint picks
  the right one per glyph.
- Title 32pt and up, body 20pt and up. Text inside a diagram, flow chart, or any
  grouped shape is artwork, not slide copy: it is exempt from the font and size
  rules, and the audit skips it. Judge it by eye in the render instead — if a label
  is unreadable from the back of the room, the figure needs redrawing, and bumping
  its point size would just burst the box it sits in.
- The page number is 28pt, pure black, fixed at the right-top corner. Never gray,
  never small: it is what the audience calls out when they want to return to a
  slide in Q&A. Cover, chapter separator and closing slides carry no number —
  nobody cites those pages. `addTitleSlide`, `addChapterSlide` and
  `addClosingSlide` use a second master without the field, so this is automatic.
- Bold marks highlighted numerals: write `**0.86**` or `**81,348 patches**` in a
  summary sentence and the kit renders that span bold (rule 7).
- No color coding beyond black text, gray captions, and red enclosures. Red is
  reserved for linkage (rule 8) so it keeps its meaning.

## Writing the summary sentences

Plain, compact, declarative. The sentence states a result, not an intention.

- Lead with the conclusion, then the condition, then the exception (rule 6).
- Include the number (rule 7). "Accuracy improved" is not a summary sentence;
  "F1 reached **0.82**, **-4.8%** relative to manual matching" is.
- 2-4 sentences per slide, hard stop at 4.
- No supplementary explanation in the summary block (rule 9). Caveats, definitions,
  and "note that..." material go in a `caption()` next to the figure or in a
  `balloon()` pointing at the part they explain.
- No hedging, no throat-clearing, no transition words. Cut "In this study, we
  propose to..." down to the claim itself.

**Example 1:**
Draft: "We conducted experiments on the benchmark dataset to evaluate the proposed
method, and the results were generally good although some cases were difficult."
Rewrite: "Trained on **1.3M** samples; IoU reached **0.79**." / "Failures
concentrate in the **dense** class (**12%** of samples)."

**Example 2:**
Draft title: "5.2 Results [2/2]"
Rewrite title: "Reconstruction accuracy by input class"

## Titles

Keyword-led and specific to that one slide (rule 4). Never a bare section number or
a generic label ("Results", "Discussion"). If a title needs a chapter prefix to make
sense, insert a chapter separator slide instead and drop the prefix (rule 5).

## Linking claims to evidence

When the sentence and the part of the figure it refers to are not obviously paired,
mark it (rule 8):

- `redBox(slide, {x, y, w, h})` / `red_box(slide, x=, y=, w=, h=)` — red enclosure
  around the row, bar, or region.
- `balloon(slide, {text, x, y, w, h, pointTo: {x, y}})` /
  `balloon(slide, text=, x=, y=, w=, h=, point_to={"x":, "y":})` — for supplementary
  text that must sit next to the figure rather than in the summary (rule 9).

One or two per slide. If a slide needs five enclosures, it is carrying two slides
worth of content.

## Scripts

| Script | Use |
|---|---|
| `scripts/slide_kit.js` | `require()` it from the generator script; provides `createDeck`, `addContentSlide`, `addChapterSlide`, `addTitleSlide`, `addClosingSlide`, `fitImages`, `redBox`, `balloon`, `caption` |
| `scripts/slide_kit.py` | same kit on `python-pptx`, snake_case: `create_deck`, `add_content_slide`, `add_chapter_slide`, `add_title_slide`, `add_closing_slide`, `fit_images`, `red_box`, `balloon`, `caption`, then `deck.save(path)` |
| `scripts/check_rules.py deck.pptx [--font-en F --font-ja F]` | Audits fonts, sizes, bullet counts, page number, margins, evidence-zone use. FAIL must be fixed; WARN needs a reason |
| `scripts/page_number.py` | `add_page_number(slide, n, slide_w)` — drops a 28pt black `slidenum` field at the right-top of an existing deck; returns False and skips cover, separator and closing slides. Use it when retrofitting a deck you did not generate |

Minimal generator:

```js
const { createDeck, addContentSlide, addChapterSlide, fitImages, redBox, caption } =
  require('./scripts/slide_kit.js');

const deck = createDeck({ fontEn: 'Calibri', fontJa: 'Yu Gothic' });
addChapterSlide(deck, 'Chapter 3: Method');
const { slide, zone } = addContentSlide(deck, {
  title: 'Dataset and study area',
  summary: [
    'Test data: **81,348** patches (1,024 px) over the **4** validation areas.',
    'Training data: **1,222** images with **186,000+** annotated instances.',
  ],
});
fitImages(slide, zone, [{ path: 'map.png', caption: 'Scope of the study area' }]);
redBox(slide, { x: 6.2, y: 4.0, w: 2.4, h: 1.1 });
deck.pres.writeFile({ fileName: 'talk.pptx' });
```

The same deck in Python:

```python
import sys; sys.path.insert(0, 'scripts')
from slide_kit import create_deck, add_content_slide, add_chapter_slide, fit_images, red_box

deck = create_deck(font_en='Calibri', font_ja='Yu Gothic')
add_chapter_slide(deck, 'Chapter 3: Method')
slide, zone = add_content_slide(deck, title='Dataset and study area', summary=[
    'Test data: **81,348** patches (1,024 px) over the **4** validation areas.',
    'Training data: **1,222** images with **186,000+** annotated instances.',
])
fit_images(slide, zone, [{'path': 'map.png', 'caption': 'Scope of the study area'}])
red_box(slide, x=6.2, y=4.0, w=2.4, h=1.1)
deck.save('talk.pptx')
```

## Reviewing an existing deck

Run `check_rules.py` first for the mechanical violations, then read the text with
`markitdown deck.pptx` and judge rules 1, 4, 6, 7, 9 by hand — those are content
rules a script cannot see. Report violations per slide with the rule number and a
concrete rewrite, not a general comment.

A review is a report. Apply only the changes that are safe under Content precedence
— adding a missing page number, or a font or size change that provably fits in the
space already there. Everything that would move, reword, or shrink existing content
goes in the report for the user to decide, with what you would change and why.
Say plainly which fixes you applied and which you left, and never present a
violation you chose not to touch as if it were resolved.
