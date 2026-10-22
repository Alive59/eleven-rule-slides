---
name: eleven-rule-report
description: Build or fix slides for a regular project report meeting (定例会, 進捗報告, progress report to a client or funder) under the eleven slide rules plus report rules for method, evaluation and visualization pages. Use instead of eleven-rule-slides for recurring report decks.
---

# Eleven-rule report slides

A report-meeting variant of **eleven-rule-slides**. The audience sees this deck every
few weeks, so what they need from each page is: what changed since last time, how it
was measured, and what came out. Everything in eleven-rule-slides still applies —
layout zones, fonts, title rules, page number on the master, content precedence,
the audit, the render check. This skill adds rules for the four page types a report
deck is made of, and where the two disagree, this skill wins.

## Before you start

1. **Load the base skill.** Read the eleven-rule-slides SKILL.md and its
   `references/eleven-rules.md`. Use its scripts (`slide_kit.js` / `slide_kit.py`,
   `check_rules.py`, `page_number.py`). If they are not installed or are missing
   `slide_kit.py`, run
   `git clone --depth 1 https://github.com/Alive59/eleven-rule-slides` and use the
   clone, as the base skill describes.
2. **Ask the font pair** as the base skill does (default Calibri / Yu Gothic), unless
   the deck being edited already has one.
3. **Write the term sheet** (next section) before any slide. Show it to the user when
   a name or a "last time" baseline is not obvious from the material.

## Term sheet: one name, one color per thing

Report decks break down when the same model is "model-A ep18" on one page,
"前回モデル" on the next and "baseline" in a table. Before drafting, fix one table:

| Display name | What it is (internal ID, data, date) | Color |
|---|---|---|
| 前回 / Last time | the model or result shown at the previous meeting | `B2241E` red |
| 今回 / This time | the current model or result | `222222` black |
| 参照 / Reference | ground truth, upper bound, or input-based reference | `7F7F7F` gray |
| additional methods | other variants, in a fixed order | `BFBFBF`, `595959`, ... |

- The display name is what appears **everywhere**: bullets, chart category labels,
  legends, table headers, panel labels. Internal IDs (run names, epochs, checkpoint
  dates) appear once, in a method or setting bullet that maps them to the display
  name, and never inside a chart.
- The same metric keeps the same name and unit on every page (e.g. always
  「水平誤差 (m)」, never 「水平誤差」 on one page and 「XY error」 on another).
- Colors are bound to the display name for the whole deck, so 前回 is red and 今回
  is black on every chart (the red-black style).
- When a deck already uses a different palette consistently, keep it (content
  precedence) and fill the Color column from it instead. If it is inconsistent across
  pages, report that and ask before recoloring.

## Red frames: only when the user asks

This overrides base rule 8 for report decks. Do not add red enclosures (`red_box`) on
your own. The link between a sentence and its evidence is carried by the term-sheet
names, the order of the bullets, and bold values in tables. Add a red frame only when
the user asks for one and says where ("the >256 row", "the changed module"), and
only there. When reviewing a report deck, a missing red frame is not a violation.

If the user asks for a frame on a chart, say that red is also the 前回 series color
there and offer a balloon instead; draw the frame if they still want it.

## Shapes: color fill only

Every shape the skill adds — diagram blocks, legend swatches, balloons, panels — has
a solid color fill and nothing else: no outline, shadow, glow, soft edge, bevel or 3D,
gradient, pattern or reflection. Shapes whose purpose is the line itself (the red
enclosure, a dashed grouping frame, arrows and connectors) have that plain line only,
no fill and no effects.

Shapes made by python-pptx or pptxgenjs are linked to the theme through `<p:style>`,
and the default theme gives them a gradient fill and a shadow even when you set a
color. Run `flatten_shapes(deck.prs)` (helpers below) once before saving: it removes
the theme link and writes the colors out explicitly. `flat_violations(prs)` lists
what is still styled; it must return nothing. Pictures, charts and tables are not
touched. On an existing deck, report styled shapes rather than flattening them
(content precedence), unless the user asks.

## Page type 1: Method

Summary bullets state the concrete setting, in this order, one item per bullet
(3-4 bullets; merge or drop what does not apply):

1. **Input and output** — data type, size or resolution, value range.
   「入力は航空写真 (40 cm) と建築物マスク、出力は 128×128 の正規化 nDSM」
2. **Modules changed since last time** — name the module and what it now does.
   Mark it as a change: start the bullet with 「変更：」/ "Changed:" or bold the module name.
3. **Training setting** when available — loss terms, training data size, key
   hyperparameters. 「Loss：座標 + 属性 + 形状（境界・角）」「学習データ **20 万棟**」
4. Internal ID mapping if a chart later uses a display name
   (「今回 = model-B ep28」) — this is the one place the ID appears.

Abbreviations are spelled out once on the slide where they first appear, in a short
parenthesis; a longer definition goes in a caption or balloon (rule 9).

**Architecture figure**

- **A given diagram is fixed.** Do not redraw, restyle, recolor, relabel or rearrange
  it. Allowed: uniform scaling to fit the evidence zone, and a red frame around the
  changed module when the user asks for one (see *Red frames*).
- **No diagram given:** draw the simplest flow that shows the setting.
  - One row, left to right: inputs → model → output. Optional second row in a dashed
    box for how the output is used downstream.
  - Inputs and outputs as small thumbnails (about 1.1-1.5 in square) with a short
    noun label directly below.
  - The model as one rounded rectangle with one light fill (e.g. `E2EFDA`), no
    border; inside it at most 3-4 sub-blocks named with their common names.
  - Arrows plain dark gray lines. No icons. Shapes follow *Shapes* below.
  - At most about 7 boxes in total. If more are needed, split into two slides.
  - Group the shapes so the audit treats the diagram as artwork.

## Page type 2: Evaluation

Summary bullets, in this order (2-4 bullets):

1. **How the metric is calculated** — the unit of evaluation (per building, per roof
   face, per pixel), what is compared with what, the unit, and how it is aggregated.
   「頂点誤差：生成モデルと参照モデルの対応頂点間の距離 (m) の建築物平均」.
   In this skill the metric definition belongs in the summary, not in a balloon —
   this overrides the base rule 9 for this one sentence.
2. **Data** — area or dataset, number of samples, input condition.
3. **Result as a comparison** with numbers in bold: 「水平誤差 **1.00 → 0.77 m**、
   品質スコア **0.70 → 0.72**」. Before/after or last-time/this-time comes first whenever
   both exist; a single number without a baseline is the fallback.
4. Exception or remaining problem, with its share or value, if it matters.

**Charts**

- Build charts as native PowerPoint charts from the numbers (helpers below), not as
  pasted matplotlib PNGs. Ask for the numbers if only an image exists; if the user
  has none, crop the image to the plot area and say so.
- Nothing inside the chart but bars, axes, gridlines and value labels: no chart
  title, subtitle, axis title, legend, annotation, or arrow text. The metric name
  goes under the chart as a 20pt bold label (same style as visualization-page panel
  labels). A shared legend, if needed, is one row of color swatches + display names
  on the slide, once per slide.
- Category labels and series names are the term-sheet display names, in term-sheet
  order (前回 before 今回), and colors come from the term sheet.
- Red-black style, in two forms:
  - **One metric, 2-3 terms** (`compare_bars`): one column per term, left to right in
    term-sheet order; bold 16pt value labels above the columns; bold category labels;
    gray value axis (14pt, `595959`) with light gridlines (`E6E6E6`). Up to three of
    these side by side, one metric each, each with its label underneath.
  - **One metric split by bins** (`grouped_bars`): terms as adjacent columns per bin,
    value axis and gridlines hidden, 14pt value labels on the columns, one shared
    slide legend (`slide_legend`) under the chart.
- Do not show a second chart that repeats numbers the table already shows.
- Axis and data-label text 14pt or larger. Value axes start at zero by default; a
  higher minimum (`y_min`) is acceptable when the difference would otherwise be
  invisible, since every column carries its value label.

**Image grids (qualitative comparison inside an evaluation page)**

- Pure images only: crop away any titles, axis ticks, colorbars, or text drawn into
  the image by the plotting script.
- Column labels (「画像」「正解」「今回」) are slide text boxes centered under each
  column, using the display names.

**Tables**

Keep the style of a clean comparison table (`report_table()` below):

- White background, no fills, no banding. Text 14pt, black.
- Two header rows, bold: the group (前回 / 今回, or area names) merged over its
  metric columns, then the metric names with units.
- Row labels left-aligned and bold; values centered, same number of decimals per column.
- Rules: 1pt black vertical line after the row-label column and between groups;
  1pt `595959` line under the header and under the last row; 0.75pt `D9D9D9` lines
  between body rows. No other borders.
- The total row (全体 / ALL) is last.
- Bold the better value of each pair. A red frame around a row only when the user
  asks for it (see *Red frames*).

## Page type 3: Visualization

Use the four-panel layout: one row of equal panels across the evidence zone.

| Element | Geometry (in) |
|---|---|
| Panels (4) | y ≈ 3.4, width ≈ 2.8-3.0, equal height, even gaps across x = 0.6-12.9 |
| Colorbar or category legend | directly above the panel it belongs to, y ≈ 2.65-3.25 |
| Panel labels | 20pt bold, centered under each panel, y ≈ 6.65 |

- Fixed panel order on every visualization page, for example:
  [summary metric chart] [per-item metric map + colorbar] [input image] [result +
  category legend]. Use three panels if one does not exist; do not reorder.
- The summary chart in the first panel follows the evaluation-page chart rules
  (native, no inner title, red-black term-sheet colors). Every area uses the same
  series colors; areas are told apart by the slide title, not by chart color.
- Colorbars use the same range and the same colormap on every page, so two areas can
  be compared by eye.

Summary bullets give the concrete setting of what is shown, with the same fields in
the same order on every visualization page:

1. **Data** — source and resolution of the input, and what kind of input it is
   (reference data vs extracted). 「航空写真 (40 cm) から抽出したフットプリントを入力」
2. **Scale and headline value** — number of items in the area and the main metric.
   「建築物 **12,000 棟**、品質スコア **0.77**」. Add the number shown in the panels if it
   is a subset (「図はブロック内 **90 棟**」).
3. **Observation** — the result in one sentence, with a number when possible.

If a field is unknown for one page, ask; do not drop it on that page only.

## Page type 4: Progress overview and next steps

- Progress: what was done since the last meeting, as 2-4 bullets with numbers
  (「学習データ 10 → 20 → **50 万棟**」).
- Next steps: what is in progress and what is planned, one bullet each, noun-ending.
- No nested bullets deeper than one level.

## Language (Japanese and English)

Plain words, formal expression. Write the way a careful report to a client reads,
not the way a chat or a paper abstract reads.

- Use common words. Avoid words that are rare in ordinary work communication,
  internal jargon, and slang.
- Japanese: one consistent ending per deck — noun-ending (体言止め) or 「〜した／〜である」.
  No です・ます in bullets. Keep particles; no colloquial contraction.
- English: short declarative sentences, active voice, no hype words.
- Define each abbreviation once per deck, at first use.
- Numbers with units and thousands separators (12,000 棟; 0.77 m).

| Avoid | Use |
|---|---|
| 形状整ってないため | 形状が整っていないため |
| 一部のみLoD2あるのため | 一部の地域にのみ LoD2 があるため |
| 推測・抜粋されたもの | 推定したもの |
| アブレーション | 要素ごとの効果検証 |
| ep18 のスクラッチ版 | 前回（term sheet の表示名） |
| utilize / leverage | use |
| elucidate / showcase | explain / show |
| a plethora of | many |
| SOTA (undefined) | state-of-the-art |

## Workflow

1. Base-skill steps 0-2 (fonts, collect content, outline), with the term sheet added
   to step 1 and each page tagged as method / evaluation / visualization / overview.
2. Generate with the base kit (`add_content_slide` etc.), then add charts and tables
   with the helpers below on the returned slide.
3. Audit with `check_rules.py` (charts and tables are evidence and are not
   font-checked there).
4. `flatten_shapes(deck.prs)` before `deck.save()`, then `flat_violations()` on the
   saved file must return an empty list.
5. Report-rule check by hand, per page: term-sheet names and colors used everywhere;
   no text inside charts; metric definition bullet on every evaluation page; same
   bullet fields on every visualization page; language table.
6. Render and look at every page, as in the base skill.

When reviewing an existing report deck, the base skill's content precedence holds:
report violations of these rules per page with a concrete rewrite, and apply only the
safe ones. Recoloring charts, rebuilding image charts as native charts, and renaming
series are changes to existing content — propose them, apply them when the user agrees.

## Helpers (python-pptx)

The same code is in the repo at `eleven-rule-report/scripts/report_kit.py`
(github.com/Alive59/eleven-rule-slides, branch `eleven-rule-report`), with a smoke
test in `eleven-rule-report/tests/`. If that file is not available, save the block
below as `report_kit.py` next to the generator script. Tested with python-pptx 1.0.2.
Set `PALETTE` from the term sheet.

```python
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
import copy
from lxml import etree

# Red-black style. Colors come from the term sheet, bound to display names.
PALETTE = {"前回": "B2241E", "今回": "222222", "参照": "7F7F7F"}
GRID, AXIS_LINE, AXIS_TEXT = "E6E6E6", "BFBFBF", "595959"
AXIS_PT, LABEL_PT = 14, 16

def _ea(font, face):
    font._rPr.append(etree.Element(qn("a:ea"), typeface=face))

def _style_common(chart, show_value_axis=True, font_en="Calibri", font_ja="Yu Gothic"):
    chart.has_title = False
    chart.has_legend = False
    chart.font.size = Pt(AXIS_PT); chart.font.name = font_en; _ea(chart.font, font_ja)
    ca, va = chart.category_axis, chart.value_axis
    ca.has_title = False; va.has_title = False
    ca.format.line.color.rgb = RGBColor.from_string(AXIS_LINE)
    ca.tick_labels.font.bold = True; ca.tick_labels.font.size = Pt(AXIS_PT)
    ca.tick_labels.font.color.rgb = RGBColor(0, 0, 0)
    if show_value_axis:
        va.has_major_gridlines = True
        va.major_gridlines.format.line.color.rgb = RGBColor.from_string(GRID)
        va.format.line.color.rgb = RGBColor.from_string(AXIS_LINE)
        va.tick_labels.font.size = Pt(AXIS_PT)
        va.tick_labels.font.color.rgb = RGBColor.from_string(AXIS_TEXT)
        va.tick_labels.number_format = "General"; va.tick_labels.number_format_is_linked = False
    else:
        va.visible = False; va.has_major_gridlines = False
    plot = chart.plots[0]
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.font.size = Pt(LABEL_PT); dl.font.bold = True; dl.font.color.rgb = RGBColor(0, 0, 0)
    dl.number_format_is_linked = True
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    return plot

def compare_bars(slide, x, y, w, h, labels, values, number_format="0.0", y_min=None):
    """One metric, one column per term (前回/今回/参照...), left to right. Colors by label."""
    cd = CategoryChartData(number_format=number_format)
    cd.categories = labels
    cd.add_series("v", values)
    chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(x), Inches(y),
                                   Inches(w), Inches(h), cd).chart
    plot = _style_common(chart); plot.gap_width = 50
    if y_min is not None:
        chart.value_axis.minimum_scale = y_min
    for i, lab in enumerate(labels):
        f = plot.series[0].points[i].format.fill
        f.solid(); f.fore_color.rgb = RGBColor.from_string(PALETTE[lab])
    return chart

def grouped_bars(slide, x, y, w, h, categories, series, number_format="0.0"):
    """Terms as series per category (e.g. bins), value axis hidden, labels on bars.
    series = {"前回": [...], "今回": [...]}. Pair with slide_legend() once per slide."""
    cd = CategoryChartData(number_format=number_format)
    cd.categories = categories
    for name, vals in series.items():
        cd.add_series(name, vals)
    chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(x), Inches(y),
                                   Inches(w), Inches(h), cd).chart
    plot = _style_common(chart, show_value_axis=False); plot.gap_width = 60; plot.overlap = 0
    plot.data_labels.font.size = Pt(AXIS_PT); plot.data_labels.font.bold = False
    for s, name in zip(plot.series, series):
        s.format.fill.solid(); s.format.fill.fore_color.rgb = RGBColor.from_string(PALETTE[name])
    return chart

def slide_legend(slide, x, y, names, font_en="Calibri", font_ja="Yu Gothic", size=20):
    """One shared legend row on the slide: swatch + display name per term."""
    from pptx.enum.shapes import MSO_SHAPE
    cx = x
    for name in names:
        sw = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx), Inches(y + 0.1), Inches(0.6), Inches(0.22))
        plain_fill(sw, PALETTE[name])
        tb = slide.shapes.add_textbox(Inches(cx + 0.7), Inches(y - 0.05), Inches(1.6), Inches(0.5))
        r = tb.text_frame.paragraphs[0].add_run(); r.text = name
        r.font.size = Pt(size); r.font.name = font_en; _ea(r.font, font_ja)
        cx += 2.4

def label_under(slide, x, w, text, y=6.65, font_en="Calibri", font_ja="Yu Gothic", size=20):
    """20pt bold label centered under a chart or panel (replaces in-chart titles)."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(0.45))
    p = tb.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = text; r.font.size = Pt(size); r.font.bold = True
    r.font.name = font_en; _ea(r.font, font_ja)

def _set_borders(cell, spec):
    """spec = {"L": (pt, "RRGGBB") or None, ...}; written in schema order, before the fill."""
    tcPr = cell._tc.get_or_add_tcPr()
    for side in "LRTB":
        for old in tcPr.findall(qn(f"a:ln{side}")):
            tcPr.remove(old)
    for i, side in enumerate("LRTB"):
        w_pt, color = spec.get(side) or (0, None)
        ln = etree.Element(qn(f"a:ln{side}"), w=str(int(w_pt * 12700)))
        if w_pt:
            etree.SubElement(etree.SubElement(ln, qn("a:solidFill")), qn("a:srgbClr"), val=color)
        else:
            etree.SubElement(ln, qn("a:noFill"))
        tcPr.insert(i, ln)

def report_table(slide, x, y, w, h, row_label, groups, metrics, rows, bold=(),
                 font_en="Calibri", font_ja="Yu Gothic", size=14):
    """Page-9 style comparison table: white, no fills, bold 2-row header (group over metric),
    black 1pt vertical rule after the label column and between groups, 1pt #595959 rule under
    the header and the last (全体/ALL) row, 0.75pt #D9D9D9 rules between body rows.
    rows = [("≤16", [g1m1, g1m2, ..., g2m1, ...]), ..., ("全体", [...])]
    bold = {(body_row, value_col)} — the better value of each pair."""
    nm = len(metrics); ncol = 1 + nm * len(groups); nrow = 2 + len(rows)
    gf = slide.shapes.add_table(nrow, ncol, Inches(x), Inches(y), Inches(w), Inches(h))
    tbl = gf.table; tbl.first_row = False; tbl.horz_banding = False
    def put(r, c, text, b=False, align=PP_ALIGN.CENTER):
        cell = tbl.cell(r, c); cell.fill.background(); cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = cell.text_frame.paragraphs[0]; p.alignment = align
        if text != "":
            run = p.add_run(); run.text = str(text)
            run.font.size = Pt(size); run.font.bold = b; run.font.name = font_en
            run.font._rPr.append(etree.Element(qn("a:ea"), typeface=font_ja))
    put(0, 0, ""); put(1, 0, row_label, True, PP_ALIGN.LEFT)
    for g, name in enumerate(groups):
        c0 = 1 + g * nm
        tbl.cell(0, c0).merge(tbl.cell(0, c0 + nm - 1)); put(0, c0, name, True)
        for m, met in enumerate(metrics):
            put(1, c0 + m, met, True)
    for r, (lab, vals) in enumerate(rows):
        put(2 + r, 0, lab, True, PP_ALIGN.LEFT)
        for c, v in enumerate(vals):
            put(2 + r, 1 + c, v, (r, c) in set(bold))
    for r in range(nrow):
        for c in range(ncol):
            spec = {}
            if c == 0 or (c % nm == 0 and c != ncol - 1):
                spec["R"] = (1.0, "000000")
            if c > 1 and (c - 1) % nm == 0:          # merged group headers need the left rule too
                spec["L"] = (1.0, "000000")
            if r == 1 or r == nrow - 1:
                spec["B"] = (1.0, "595959")
            elif r >= 2:
                spec["B"] = (0.75, "D9D9D9")
            _set_borders(tbl.cell(r, c), spec)
    return tbl

# --- Flat shapes: color fill only -------------------------------------------------
# Autoshapes carry a <p:style> that links them to the theme: fillRef idx=3 is a
# gradient and effectRef idx=2 a shadow in the default theme. An empty effectLst is
# not enough (some renderers still draw the theme effect), so the link is removed and
# its colors are written out explicitly.

_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
_FILLS = {_A + t for t in ("noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill")}
_COLORS = {_A + t for t in ("schemeClr", "srgbClr", "prstClr", "sysClr", "scrgbClr", "hslClr")}


def _color_of(ref):
    """Copy of the color element inside a style ref (fillRef/lnRef/fontRef), or None."""
    if ref is None:
        return None
    for c in ref:
        if c.tag in _COLORS:
            return copy.deepcopy(c)
    return None


def _solid(color_el):
    f = etree.Element(_A + "solidFill")
    f.append(color_el)
    return f


def _flatten(el):
    """el is a p:sp or p:cxnSp. Leaves: solid fill (filled shapes, no outline) or a
    plain line (unfilled shapes and connectors). No style link, effects, 3D, gradient."""
    spPr = el.find(_P + "spPr")
    style = el.find(_P + "style")
    is_cxn = etree.QName(el).localname == "cxnSp"
    fill_ref = _color_of(style.find(_A + "fillRef")) if style is not None else None
    ln_ref = _color_of(style.find(_A + "lnRef")) if style is not None else None
    font_ref = _color_of(style.find(_A + "fontRef")) if style is not None else None

    for tag in ("effectLst", "effectDag", "scene3d", "sp3d"):
        for x in spPr.findall(_A + tag):
            spPr.remove(x)
    fills = [c for c in spPr if c.tag in _FILLS]
    fill = fills[0] if fills else None
    if fill is not None and fill.tag in (_A + "gradFill", _A + "pattFill"):
        stop = fill.find(".//" + _A + "gs")       # first gradient stop -> solid
        col = next((copy.deepcopy(c) for c in (stop if stop is not None else fill)
                    if c.tag in _COLORS), None)
        new = _solid(col) if col is not None else etree.Element(_A + "noFill")
        fill.addprevious(new); spPr.remove(fill); fill = new
    if fill is None:
        if is_cxn or style is None or fill_ref is None:
            fill = etree.Element(_A + "noFill")
        else:
            fill = _solid(fill_ref)
        geom = [c for c in spPr if c.tag in (_A + "xfrm", _A + "prstGeom", _A + "custGeom")]
        (geom[-1].addnext(fill) if geom else spPr.insert(0, fill))
    filled = fill.tag != _A + "noFill"

    ln = spPr.find(_A + "ln")
    if ln is None:
        ln = etree.Element(_A + "ln")
    for x in ln.findall(_A + "effectLst") + ln.findall(_A + "effectDag"):
        ln.remove(x)
    if filled and not is_cxn:                     # filled shape: no outline at all
        for child in list(ln):
            ln.remove(child)
        etree.SubElement(ln, _A + "noFill")
    elif not any(c.tag in _FILLS for c in ln):    # line shape relying on the theme color
        ln.insert(0, _solid(ln_ref if ln_ref is not None else etree.Element(_A + "srgbClr", val="000000")))
    fill.addnext(ln)
    ln.addnext(etree.Element(_A + "effectLst"))

    if style is not None:
        if font_ref is not None:                  # keep the text color the style gave
            for rpr in el.iter(_A + "rPr", _A + "endParaRPr"):
                if not any(c.tag in _FILLS for c in rpr):
                    rpr.insert(0, _solid(copy.deepcopy(font_ref)))
            for r in el.iter(_A + "r"):
                if r.find(_A + "rPr") is None:
                    rpr = etree.Element(_A + "rPr"); r.insert(0, rpr)
                    rpr.append(_solid(copy.deepcopy(font_ref)))
        el.remove(style)


def plain_fill(shape, color):
    """Solid color fill and nothing else: no outline, shadow, glow, bevel or gradient."""
    shape.fill.solid(); shape.fill.fore_color.rgb = RGBColor.from_string(color)
    _flatten(shape._element)
    return shape


def _shape_elements(prs):
    for i, slide in enumerate(prs.slides, 1):
        for el in slide.shapes._spTree.iter(_P + "sp", _P + "cxnSp"):
            if el.find(".//" + _P + "cNvSpPr[@txBox='1']") is not None:
                continue                          # text boxes carry no shape styling
            if el.find(_P + "spPr") is None:
                continue
            yield i, el


def flatten_shapes(prs):
    """Pass over a whole deck before saving: every autoshape and connector ends with a
    color fill only (filled shapes) or a plain line only (red enclosures, dashed frames,
    arrows). Pictures, charts, tables and text boxes are left alone. Returns the count."""
    n = 0
    for _, el in _shape_elements(prs):
        _flatten(el); n += 1
    return n


def flat_violations(prs):
    """[(slide_no, shape_name, reason)] for shapes styled beyond a color fill."""
    out = []
    for i, el in _shape_elements(prs):
        name = el.find(".//" + _P + "cNvPr").get("name")
        spPr = el.find(_P + "spPr")
        if el.find(_P + "style") is not None:
            out.append((i, name, "linked to theme styles (shadow/gradient may apply)"))
        for tag in ("scene3d", "sp3d", "gradFill", "pattFill"):
            if spPr.find(_A + tag) is not None:
                out.append((i, name, f"has {tag}"))
        eff = spPr.find(_A + "effectLst")
        if (eff is not None and len(eff)) or spPr.find(_A + "effectDag") is not None:
            out.append((i, name, "has shadow/glow/soft edge"))
        fills = [c for c in spPr if c.tag in _FILLS]
        ln = spPr.find(_A + "ln")
        if fills and fills[0].tag == _A + "solidFill" and etree.QName(el).localname == "sp" \
                and ln is not None and ln.find(_A + "noFill") is None:
            out.append((i, name, "filled shape has an outline"))
    return out
```

Usage on a slide from the base kit:

```python
slide, zone = add_content_slide(deck, title="評価対象全体の水密性", summary=[
    "水密性：生成モデルが閉じた面で構成される建築物の割合 (%)",
    "評価対象 **2,500 棟**、面数で 5 区分",
    "水密性 **78.2 → 84.5%**",
])
compare_bars(slide, 0.6, 2.9, 3.8, 3.6, ["前回", "今回"], [78.2, 84.5], y_min=60)
label_under(slide, 0.6, 3.8, "水密性 (%)")
grouped_bars(slide, 4.8, 2.9, 7.9, 3.4, ["≤16", "17-32", "ALL"],
             {"前回": [100, 99.3, 77.7], "今回": [100, 99.3, 84.1]})
slide_legend(slide, 6.8, 6.5, ["前回", "今回"])
flatten_shapes(deck.prs)          # last step before saving: color fill only
deck.save("report.pptx")
```

The JS kit users can do the same with `pptxgenjs` `slide.addChart` / `slide.addTable`:
`showTitle: false`, `showLegend: false`, `showValue: true`, `dataLabelFontBold: true`,
`chartColors` from the term sheet in category order, `valGridLine: { color: "E6E6E6" }`
(or `valAxisHidden: true` for grouped bins), and table borders as specified above.
Give every shape `shadow: { type: "none" }` and either `line: { type: "none" }` (filled)
or `fill: { type: "none" }` (line shapes), then run `flatten_shapes()` from Python on
the saved file, since pptxgenjs also writes the theme link.
