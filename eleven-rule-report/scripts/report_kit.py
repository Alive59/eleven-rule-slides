"""Report-page helpers for eleven-rule-report (python-pptx).

Native charts in the red-black style and the comparison table style, to be used on
slides made by ../../scripts/slide_kit.py. Set PALETTE from the deck's term sheet.
"""
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
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
        sw.fill.solid(); sw.fill.fore_color.rgb = RGBColor.from_string(PALETTE[name]); sw.line.fill.background(); sw.shadow.inherit = False
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
