---
name: eleven-rule-slides
description: Build research presentation slides that follow the eleven rules of slide composition — title + 2-4 summary sentences + supporting evidence per slide, chapter separator slides, red enclosures linking claims to figures, large fonts (title 32pt and up, body 20pt and up), Calibri for English and Yu Gothic for Japanese, page number on the slide master (top-right, none on the title and acknowledgement slides). Use this skill whenever the user asks for slides, a deck, a presentation, a talk, a progress report, a seminar/lab meeting/conference/defense presentation, or a .pptx file, and also when they ask to review, restructure, or fix existing slides. Applies even if they don't mention the eleven rules by name.
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

## Script freshness

The canonical scripts live at https://github.com/Alive59/eleven-rule-slides. Before
generating or auditing, check that this skill's `scripts/` folder has
`slide_kit.py` and that `scripts/page_number.py` defines
`install_master_page_number`. If either is missing, the bundled scripts are older
than this file: run `git clone --depth 1 https://github.com/Alive59/eleven-rule-slides`
into the working directory and use that clone's `scripts/` and `references/`
instead. If the clone fails (no network), use the bundled scripts and tell the user
that the page number and title checks will follow the older rules.

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

Both kits put the page number on the slide master and give the title and
acknowledgement slides a layout that hides it, so the two produce the same
package: zone geometry, the two-font split, minimum sizes and which slides carry
a number are identical.

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
   **Template:** if the user gave a template (.potx, or a deck whose look they
   want), it becomes the background; follow *Custom template* below. If a template
   names its own fonts, offer that pair as the default in the font question.
1. **Collect the content before touching code.** For each planned slide write, in
   plain text: the title (a compact keyword phrase, not a sentence), the 2-4
   summary sentences, and which figure or table is the evidence. If a summary
   sentence has no evidence, or a figure has no sentence, the slide is not ready.
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

## Custom template

When the user provides a template, its look wins: master and layout backgrounds,
bands, logos and decorative art stay exactly as designed. The eleven-rule layer goes
on top unchanged: the four zones, title and body sizes, the font pair, the page
number (28pt black, right-top, on the master), title and summary rules, red only
for linkage. Do not restyle, recolor or delete anything in the template.

1. **Python kit only.** `create_deck(template=path)`; pptxgenjs cannot open a
   template.
2. **Check the slide size first.** The kit sets 13.333 x 7.5 in, but template art
   does not rescale with the page. If the template is another size (4:3, A4), tell
   the user and ask: convert the template to 16:9, or scale the zones to its size.
3. **Drop the template's sample slides** when it is a template (.potx, or slides
   that only hold placeholder text). When it is a real deck, ask whether to append to
   it or only borrow its masters (content precedence).
4. **Pick the layout that carries the design.** The kit defaults to the layout named
   "Blank", which in many templates has no band or logo. Render one slide per layout
   (or read the layout names) and choose the content layout, usually "Title Only".
   For the title and acknowledgement slides use the template's own title layout,
   with `showMasterSp="0"` set on it so the master's page number does not print.
   Set the attribute on the layout itself; do not use `plain_layout()` to copy a
   template layout — it copies the XML but not the image relationships, so pictures
   in the layout break.
5. **Check collisions in the render.** If template art overlaps the title zone,
   the page-number corner, or the evidence zone, keep the zones and report it; offer
   the smallest shift (e.g. title moved below a header band) and apply it only if
   the user agrees. If the template already has a slide-number field, the kit reuses
   it as is — restyle it to 28pt black right-top, and tell the user.
6. **Colors.** Text stays black on light backgrounds. On a dark template
   background, ask which text color to use before generating.

```python
from page_number import install_master_page_number

def use_template(deck, content_layout, title_layout=None):
    prs = deck.prs
    for sid in list(prs.slides._sldIdLst):          # drop the template's sample slides
        prs.part.drop_rel(sid.rId); prs.slides._sldIdLst.remove(sid)
    lay = prs.slide_layouts.get_by_name(content_layout)
    lay._element.attrib.pop("showMasterSp", None)  # master number must show
    install_master_page_number(lay.slide_master, 13.333, font=deck.font_en)
    deck.layout = lay
    if title_layout:
        tl = prs.slide_layouts.get_by_name(title_layout)
        tl._element.set("showMasterSp", "0")         # no number; layout art stays
        deck.plain_layout = tl

deck = create_deck(font_en="Calibri", font_ja="Yu Gothic", template="lab.potx")
use_template(deck, "Title Only", "Title Slide")
```

`check_rules.py` must still pass on the result; a FAIL on the title slide's page
number means the title layout still shows master shapes.

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
  Both kits split mixed strings run-by-run automatically — pass one string and it
  handles the switch, writing the Latin face to `<a:latin>` and the Japanese face
  to `<a:ea>` so PowerPoint picks the right one per glyph.
- Title 32pt and up, body 20pt and up. Text inside a diagram, flow chart, or any
  grouped shape is artwork, not slide copy: it is exempt from the font and size
  rules, and the audit skips it. Judge it by eye in the render instead — if a label
  is unreadable from the back of the room, the figure needs redrawing, and bumping
  its point size would just burst the box it sits in.
- The page number is 28pt, pure black, fixed at the right-top corner. Never gray,
  never small: it is what the audience calls out when they want to return to a
  slide in Q&A. It lives **on the slide master**, once, as a `slidenum` field —
  not as an object on each slide — so every slide prints its own number, and a
  slide added later in PowerPoint is numbered without anyone remembering to.
- Every slide is numbered **except the title slide and the acknowledgement
  (closing) slide**. Chapter separators are numbered like any other slide.
  `addTitleSlide` / `addClosingSlide` (`add_title_slide` / `add_closing_slide`) put
  those two on a plain layout that hides master shapes (`showMasterSp="0"`), so
  this is automatic.
- Bold marks highlighted numerals: write `**0.86**` or `**81,348 patches**` in a
  summary sentence and the kit renders that span bold (rule 7).
- No color coding beyond black text, gray captions, and red enclosures. Red is
  reserved for linkage (rule 8) so it keeps its meaning. Vector sketches use the
  gray palette in *Visuals*; red inside a sketch marks the defect the summary
  sentence names, which is linkage too.

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

A title is a **compact noun phrase of descriptive words, never a sentence.** It names
what the slide is about; the claim about it belongs in the summary sentences.

- No verb, no subject, no ending punctuation. "Accuracy by input class", not
  "Accuracy improves with LoD2 input." In Japanese, end on a noun (「入力クラス別の
  再構成精度」), not a predicate (「LoD2入力で精度が向上した」).
- Aim for 2-6 words (about 8 at most), or 20 Japanese characters at most.
- If the draft title is a finding, move it into the first summary sentence and keep
  only its subject as the title.

The kits reject a title that ends in sentence punctuation, ends on a Japanese
predicate, or contains a finite verb or "we/our"; `check_rules.py` fails the first
two and warns on the rest and on long titles.

**Example 3:**
Draft title: "The proposed method outperforms baselines on dense areas"
Rewrite title: "Comparison with baselines in dense areas"
(The finding, with its number, goes in the first summary sentence.)

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

## Visuals: vector first

Diagrams, schematics, building sketches, cross-sections, flows, legends and marks
are drawn from code as vectors. Never paste a screenshot of a diagram, a raster
export, or an AI-generated picture in their place. In order of preference:

1. **SVG written by code** — a Python string builder or a JS module that returns
   SVG markup — placed with `add_svg()` below (vector `svgBlip` plus a PNG fallback
   for old viewers). Use it for anything with many parts: building sketches,
   DSM profiles, before/after panels. One SVG per evidence panel.
2. **Native shapes** (`build_freeform` in python-pptx, `custGeom` / shapes in
   pptxgenjs), grouped. Use when the user will edit labels or boxes in PowerPoint,
   or for simple box-and-arrow flows.

Raster (PNG/JPG) only for content that is an image: photos, aerial imagery,
renders, result maps, screenshots of software.

Write the drawing as reusable functions (a sketch module): one function per object,
with its state as a parameter — `gable_building(..., missing=(i, n))`, a flat-roof
building, a DSM profile, a footprint bar, check/cross marks. The same object then
looks the same on every slide and in both halves of a before/after comparison.
Change the parameter; do not redraw.

**Style** (reference: building sketches in an AFD/SFD comparison deck):

- Oblique view; each face one flat fill by orientation: front `D9D9D9`, side
  `BFBFBF`, roof `8C8C8C`, back roof `A6A6A6`. White 1.5-unit strokes separate the
  faces. No gradients, shadows, or dark outlines.
- Defects: a red (`B2241E`) circle around the spot; a missing part as a white
  polygon with a red dashed outline. A correct result carries no red.
- Lines: data profiles `222222`, 3 units, round joins; ground and reference lines
  `7F7F7F`, 2 units; arrows `595959` with one arrowhead.
- Panels: light rounded rectangles (`F2F2F2` / `E7E6E6`), no border. A category
  color (e.g. light blue for AI estimate, light orange for external data) only
  when a legend on the slide defines it.
- Result lists beside a sketch: ✗ in red for failures, ✓ in black for successes.
- Text in the SVG uses the deck's faces (Japanese face first in the stack). Set
  the viewBox to 100 units per inch of placed width, so font-size 20 = 14.4pt;
  keep labels at 20 units or more. SVG text is artwork (exempt from the audit), but
  it must be readable in the render.
- Fill the evidence zone: size the viewBox to the zone's aspect (12.1 x 4.35 in →
  1210 x 435) rather than placing a small drawing in a large empty area.

```python
import os, subprocess
from lxml import etree
from pptx.util import Inches
from pptx.oxml.ns import qn
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

SVG_NS = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
FONT = "'Yu Gothic','YuGothic','Noto Sans CJK JP','Hiragino Sans',sans-serif"
FACE = {"front": "#D9D9D9", "side": "#BFBFBF", "roof": "#8C8C8C", "roof_back": "#A6A6A6"}
RED, INK = "#B2241E", "#222222"

def add_svg(slide, svg_path, x, y, w, h=None):
    """SVG as a vector picture with a PNG fallback; aspect kept inside w x h."""
    vb = [float(v) for v in etree.parse(svg_path).getroot().get("viewBox").split()]
    a = vb[3] / vb[2]
    if h is None or h / w > a: h = w * a
    else: w = h / a
    png = os.path.splitext(svg_path)[0] + "_fb.png"
    subprocess.run(["rsvg-convert", "-w", str(int(w * 200)), "-o", png, svg_path], check=True)
    pic = slide.shapes.add_picture(png, Inches(x), Inches(y), Inches(w), Inches(h))
    pkg = slide.part.package
    part = Part(PackURI(str(pkg.next_partname("/ppt/media/image%d.svg"))),
                "image/svg+xml", pkg, open(svg_path, "rb").read())
    ext = etree.SubElement(etree.SubElement(pic._element.find(".//" + qn("a:blip")),
          qn("a:extLst")), qn("a:ext"), uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}")
    etree.SubElement(ext, "{%s}svgBlip" % SVG_NS, nsmap={"asvg": SVG_NS}).set(
        qn("r:embed"), slide.part.relate_to(part, RT.IMAGE))
    return pic

def _poly(pts, fill, stroke="#FFFFFF", sw=1.5, dash=None):
    d = f' stroke-dasharray="{dash}"' if dash else ""
    p = " ".join(f"{a:.1f},{b:.1f}" for a, b in pts)
    return f'<polygon points="{p}" fill="{fill}" stroke="{stroke}" stroke-width="{sw}"{d}/>'

def gable_building(x, y, w=165, wall=52, ridge=61, depth=40, slope=34, missing=None):
    """Oblique gable-roof building; front face bottom-left at (x, y), y = ground.
    missing=(i, n): roof in n strips, strip i a red dashed gap with a red circle."""
    top, ry, h = y - wall, y - wall - ridge, depth / 2
    back = ry + 0.8 * slope
    out = [_poly([(x+depth, back), (x+w+depth, back), (x+w+h, ry), (x+h, ry)], FACE["roof_back"]),
           _poly([(x+w, y), (x+w+depth, y-slope), (x+w+depth, back), (x+w+h, ry), (x+w, top)], FACE["side"]),
           _poly([(x, y), (x+w, y), (x+w, top), (x, top)], FACE["front"])]
    n = missing[1] if missing else 1
    s = w / n
    for k in range(n):
        pts = [(x+k*s, top), (x+(k+1)*s, top), (x+(k+1)*s+h, ry), (x+k*s+h, ry)]
        gap = missing and k == missing[0]
        out.append(_poly(pts, "#FFFFFF", RED, 2.5, "7,5") if gap else _poly(pts, FACE["roof"]))
    if missing:
        cx = x + (missing[0] + 0.5) * s + h
        out.append(f'<circle cx="{cx:.1f}" cy="{ry:.1f}" r="13" fill="none" stroke="{RED}" stroke-width="2.5"/>')
    return "\n".join(out)

def label(x, y, text, size=20, bold=False, color=INK, anchor="middle"):
    return (f'<text x="{x}" y="{y}" font-family="{FONT}" font-size="{size}" font-weight='
            f'"{"bold" if bold else "normal"}" fill="{color}" text-anchor="{anchor}">{text}</text>')

def svg_doc(width, height, body):   # 100 units per placed inch
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{width*2}" height="{height*2}" '
            f'viewBox="0 0 {width} {height}">{body}</svg>')

# usage: before/after sketch filling the evidence zone
body = (gable_building(150, 330, missing=(1, 3)) + label(250, 380, "AIなし：屋根面の欠落", bold=True, color=RED)
        + gable_building(750, 330) + label(850, 380, "AIあり：閉じた LoD2 形状", bold=True))
open("sketch.svg", "w").write(svg_doc(1210, 435, body))
add_svg(slide, "sketch.svg", 0.6, 2.75, 12.1, 4.35)
```

`rsvg-convert` makes the fallback (`cairosvg` works too). In pptxgenjs, write the SVG
the same way from a JS module and add it with `slide.addImage({ path: 'sketch.svg' })`,
then check the render. LibreOffice previews draw the SVG, so judge it there.

## Scripts

| Script | Use |
|---|---|
| `scripts/slide_kit.js` | `require()` it from the generator script; provides `createDeck`, `addContentSlide`, `addChapterSlide`, `addTitleSlide`, `addClosingSlide`, `fitImages`, `redBox`, `balloon`, `caption` |
| `scripts/slide_kit.py` | same kit on `python-pptx`, snake_case: `create_deck`, `add_content_slide`, `add_chapter_slide`, `add_title_slide`, `add_closing_slide`, `fit_images`, `red_box`, `balloon`, `caption`, then `deck.save(path)` |
| `scripts/check_rules.py deck.pptx [--font-en F --font-ja F]` | Audits fonts, sizes, bullet counts, page number, margins, evidence-zone use. FAIL must be fixed; WARN needs a reason |
| `scripts/page_number.py` | `number_deck_on_master(prs)` — retrofits a deck you did not generate: puts the 28pt black `slidenum` field on each slide master and hides master shapes on the title and acknowledgement slides. Returns a report; tell the user about slides in `hidden_also` (other master graphics hidden there too) and `per_slide` (old per-slide numbers left in place) |
| `scripts/slide_kit.py` `use_template(deck, content_layout, title_layout)` | applies a custom template (see *Custom template*): drops sample slides, numbers the content layout, hides the number on the template's title layout |
| `scripts/svg_kit.py` | vector visuals (see *Visuals*): `add_svg` (svgBlip + PNG fallback), `svg_doc`, `label`, `gable_building` |

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
