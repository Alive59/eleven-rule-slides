"""Drop a 28pt black slide-number field at the right-top of an existing deck.

Cover, chapter separator and closing slides are skipped — the audience never
cites those pages, and a number on them just adds noise.
"""
from pptx.util import Inches
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml
from pptx.enum.shapes import MSO_SHAPE_TYPE

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
EMU = 914400.0
PT, BOX_W, BOX_H, TOP, RIGHT_PAD = 28, 0.95, 0.55, 0.18, 0.40
TITLE_PT = 32

FLD = (
    '<a:p xmlns:a="%s">'
    '<a:pPr algn="r"/>'
    '<a:fld id="{{1D0E7A29-9A6B-4C5E-9F0A-{n:012d}}}" type="slidenum">'
    '<a:rPr lang="en-US" sz="{sz}" b="0" dirty="0">'
    '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>'
    '<a:latin typeface="{font}"/><a:cs typeface="{font}"/>'
    '</a:rPr>'
    '<a:t>{n}</a:t>'
    '</a:fld>'
    '<a:endParaRPr lang="en-US" sz="{sz}"/>'
    '</a:p>'
) % A

VISUAL = (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.LINKED_PICTURE,
          MSO_SHAPE_TYPE.TABLE, MSO_SHAPE_TYPE.CHART, MSO_SHAPE_TYPE.GROUP)


def is_cover(slide):
    """Cover, chapter separator or closing slide: only large text, no evidence."""
    if any(s.shape_type in VISUAL or getattr(s, 'has_chart', False) for s in slide.shapes):
        return False
    body = big = 0
    for sh in slide.shapes:
        if not sh.has_text_frame or not sh.text_frame.text.strip():
            continue
        for para in sh.text_frame.paragraphs:
            if not "".join(r.text for r in para.runs).strip():
                continue
            sizes = [r.font.size.pt for r in para.runs if r.font.size]
            if sizes and min(sizes) >= TITLE_PT:
                big += 1
            else:
                body += 1
    return big > 0 and body == 0


def place_page_number_field(slide, n, slide_w, font='Calibri'):
    """Write the field unconditionally, no cover heuristic. Returns the textbox.

    ``n`` only seeds the cached text; PowerPoint recomputes the field on open, so
    a deck whose slides are reordered still numbers correctly. Callers that know
    which slides are numbered -- slide_kit.py -- use this; callers retrofitting a
    finished deck use add_page_number() instead.
    """
    box = slide.shapes.add_textbox(Inches(slide_w - BOX_W - RIGHT_PAD), Inches(TOP),
                                   Inches(BOX_W), Inches(BOX_H))
    box.text_frame.word_wrap = False
    tx = box.text_frame._txBody
    for p in tx.findall(qn('a:p')):
        tx.remove(p)
    tx.append(parse_xml(FLD.format(n=n, sz=PT * 100, font=font)))
    return box


def add_page_number(slide, n, slide_w, font='Calibri'):
    """Retrofit: number this slide unless it is a cover, separator or closing slide."""
    if is_cover(slide):
        return False
    place_page_number_field(slide, n, slide_w, font)
    return True
