"""slide_kit.py -- geometry, typography and page numbering for eleven-rule decks.

python-pptx port of ``slide_kit.js``. Same zone geometry, same two-font split, same
top-right page number, same audit result under ``check_rules.py`` -- use this one
wherever Node is unavailable (ChatGPT Code Interpreter, a plain Python sandbox) or
where the surrounding pipeline is already Python.

    from slide_kit import create_deck, add_content_slide, fit_images

    deck = create_deck(font_en="Calibri", font_ja="Yu Gothic")
    slide, zone = add_content_slide(deck, title="...", summary=["..."])
    fit_images(slide, zone, [{"path": "fig.png", "caption": "..."}])
    deck.save("talk.pptx")

Layout is 13.333 x 7.5 in (the pptxgenjs LAYOUT_WIDE canvas).

Requires ``python-pptx``. ``Pillow`` is optional: without it, images are placed to
fill their cell instead of being aspect-fit, which can stretch a figure -- install
it if you can.
"""

from __future__ import annotations

import io
import re
import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

# scripts/ is a plain directory, not a package: make the sibling importable either
# way (python scripts/gen.py, or sys.path.append("scripts") from elsewhere).
sys.path.insert(0, str(Path(__file__).resolve().parent))
from page_number import place_page_number_field  # noqa: E402

try:  # optional -- only needed for "contain" image sizing
    from PIL import Image
except ImportError:  # pragma: no cover
    Image = None

__all__ = [
    "create_deck", "new_slide", "add_content_slide", "add_chapter_slide",
    "add_title_slide", "add_closing_slide", "fit_images", "red_box", "balloon",
    "caption", "to_runs", "set_fonts", "Deck", "G", "RED", "GRAY", "BLACK",
]

# Defaults; override per deck with create_deck(font_en=..., font_ja=...) once the
# user has told you which fonts they want.
EN_FONT = "Calibri"
JA_FONT = "Yu Gothic"

RED = "C00000"
GRAY = "808080"
BLACK = "000000"

# Japanese kana, CJK ideographs, CJK punctuation, fullwidth forms.
CJK_RE = re.compile(
    r"[　-〿぀-ゟ゠-ヿ㐀-䶿一-鿿＀-￯]"
)
BOLD_SPLIT_RE = re.compile(r"(\*\*[^*]+\*\*)")

G = {
    "W": 13.333,
    "H": 7.5,
    "title": {"x": 0.6, "y": 0.28, "w": 11.5, "h": 0.75, "font_size": 32},
    "summary": {"x": 0.6, "y": 1.15, "w": 12.1, "h": 1.5, "font_size": 20},
    "evidence": {"x": 0.6, "y": 2.75, "w": 12.1, "h": 4.35},
    "caption": {"font_size": 14},
    "page_num": {"x": 11.98, "y": 0.18, "w": 0.95, "h": 0.55, "font_size": 28},
}


def set_fonts(font_en=None, font_ja=None):
    """Set the deck-wide Latin and Japanese faces. Returns the pair in effect."""
    global EN_FONT, JA_FONT
    if font_en:
        EN_FONT = font_en
    if font_ja:
        JA_FONT = font_ja
    return EN_FONT, JA_FONT


class Deck:
    """A presentation plus the page-number counter the masters cannot carry.

    python-pptx cannot define a slide master, so the page number is written as a
    real ``slidenum`` field onto each numbered slide instead of onto a master.
    PowerPoint renumbers the field itself; ``n`` only seeds the cached text.
    """

    def __init__(self, prs, font_en, font_ja):
        self.prs = prs
        self.font_en = font_en
        self.font_ja = font_ja
        self.G = G

    @property
    def slide_count(self):
        return len(self.prs.slides)

    def save(self, filename):
        self.prs.save(str(filename))
        return str(filename)

    # pptxgenjs parity
    write_file = save


# --------------------------------------------------------------------------- #
# text runs
# --------------------------------------------------------------------------- #

def _set_run_faces(run, latin, ea):
    """Write <a:latin> and <a:ea> so PowerPoint picks the right face per glyph.

    ``run.font.name`` only writes <a:latin>; Japanese glyphs are resolved through
    <a:ea>, so a mixed run styled with font.name alone renders the kana in the
    Latin face. Schema order inside <a:rPr> is latin, ea, cs.
    """
    run.font.name = latin
    rPr = run._r.get_or_add_rPr()
    latin_el = rPr.find(qn("a:latin"))
    prev = latin_el
    for tag, face in (("a:ea", ea), ("a:cs", latin)):
        el = rPr.find(qn(tag))
        if el is not None:
            rPr.remove(el)
        el = parse_xml(
            '<a:%s xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'typeface="%s"/>' % (tag.split(":")[1], face)
        )
        prev.addnext(el)
        prev = el


def _split_cjk(text):
    """[(chunk, is_cjk)] -- consecutive runs of same-script characters."""
    out = []
    buf, buf_cjk = "", None
    for ch in text:
        is_cjk = bool(CJK_RE.match(ch))
        if buf_cjk is None:
            buf_cjk = is_cjk
        if is_cjk != buf_cjk:
            out.append((buf, buf_cjk))
            buf, buf_cjk = "", is_cjk
        buf += ch
    if buf:
        out.append((buf, bool(buf_cjk)))
    return out


def to_runs(paragraph, text, font_size, bold=False, color=BLACK,
            font_en=None, font_ja=None):
    """Fill ``paragraph`` with runs split by script and by ``**bold**`` span.

    Latin spans get the deck's Latin face, CJK spans the Japanese face, and a
    ``**0.86**`` span is emitted bold -- that is how highlighted numerals (rule 7)
    are written in summary sentences.
    """
    latin = font_en or EN_FONT
    ea = font_ja or JA_FONT
    text = "" if text is None else str(text)
    made = False
    for part in BOLD_SPLIT_RE.split(text):
        if not part:
            continue
        part_bold = part.startswith("**") and part.endswith("**")
        body = part[2:-2] if part_bold else part
        for chunk, is_cjk in _split_cjk(body):
            if not chunk:
                continue
            run = paragraph.add_run()
            run.text = chunk
            run.font.size = Pt(font_size)
            run.font.bold = bool(bold or part_bold)
            run.font.color.rgb = RGBColor.from_string(color)
            _set_run_faces(run, ea if is_cjk else latin, ea)
            made = True
    if not made:  # never leave a paragraph with no explicitly styled run
        run = paragraph.add_run()
        run.text = " "
        run.font.size = Pt(font_size)
        run.font.color.rgb = RGBColor.from_string(color)
        _set_run_faces(run, latin, ea)
    return paragraph


def _textbox(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP, align=None, wrap=True):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE  # never let PowerPoint shrink below the minimum
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    if align is not None:
        tf.paragraphs[0].alignment = align
    return box


def _bullet(paragraph, space_after_pt=6):
    """Round bullet + hanging indent. <a:pPr> order: spcAft, buFont, buChar."""
    paragraph.space_after = Pt(space_after_pt)
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("marL", str(Emu(Inches(0.3125)).emu))
    pPr.set("indent", str(-Emu(Inches(0.3125)).emu))
    for el in pPr.findall(qn("a:buChar")) + pPr.findall(qn("a:buFont")):
        pPr.remove(el)
    ns = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    pPr.append(parse_xml('<a:buFont %s typeface="Arial" pitchFamily="34" charset="0"/>' % ns))
    pPr.append(parse_xml('<a:buChar %s char="•"/>' % ns))


# --------------------------------------------------------------------------- #
# slides
# --------------------------------------------------------------------------- #

def create_deck(font_en=None, font_ja=None, title=None, author=None, template=None):
    """New 13.333 x 7.5 in deck. ``template`` may point at a .pptx/.potx to inherit."""
    set_fonts(font_en, font_ja)
    prs = Presentation(str(template)) if template else Presentation()
    prs.slide_width = Inches(G["W"])
    prs.slide_height = Inches(G["H"])
    if title:
        prs.core_properties.title = title
    if author:
        prs.core_properties.author = author
    return Deck(prs, EN_FONT, JA_FONT)


def _blank_layout(prs):
    """The blank layout, by name where the template has one, else the emptiest."""
    for layout in prs.slide_layouts:
        if layout.name.strip().lower() == "blank":
            return layout
    return min(prs.slide_layouts, key=lambda l: len(l.placeholders._element))


def new_slide(deck, numbered=True):
    """Blank slide; with ``numbered`` a 28pt black page-number field at right-top."""
    slide = deck.prs.slides.add_slide(_blank_layout(deck.prs))
    for shape in list(slide.shapes):  # a layout placeholder would print its prompt text
        shape._element.getparent().remove(shape._element)
    if numbered:
        place_page_number_field(slide, deck.slide_count, G["W"], font=deck.font_en)
    return slide


def add_content_slide(deck, title=None, summary=(), notes=None):
    """Title + 2-4 summary sentences + reserved evidence zone.

    Returns ``(slide, zone)``; put figures in ``zone`` with :func:`fit_images`.
    """
    if not title:
        raise ValueError("rule 4: every slide needs a keyword title")
    summary = list(summary)
    if len(summary) > 4:
        raise ValueError(
            f"rule 6: {len(summary)} summary sentences; split the slide (max 4)"
        )
    slide = new_slide(deck)

    t = G["title"]
    tf = _textbox(slide, t["x"], t["y"], t["w"], t["h"], anchor=MSO_ANCHOR.MIDDLE,
                  wrap=False).text_frame
    to_runs(tf.paragraphs[0], title, t["font_size"], bold=True)

    if summary:
        s = G["summary"]
        tf = _textbox(slide, s["x"], s["y"], s["w"], s["h"]).text_frame
        for i, line in enumerate(summary):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            _bullet(para)
            to_runs(para, line, s["font_size"])

    if notes:
        slide.notes_slide.notes_text_frame.text = notes
    return slide, dict(G["evidence"])


def _banner_slide(deck, text, size=40, y=2.8, h=1.9):
    slide = new_slide(deck, numbered=False)
    tf = _textbox(slide, 1.0, y, 11.3, h, anchor=MSO_ANCHOR.MIDDLE,
                  align=PP_ALIGN.CENTER).text_frame
    to_runs(tf.paragraphs[0], text, size, bold=True)
    return slide


def add_chapter_slide(deck, title):
    """Chapter separator (rule 5): title only, centered, no page number."""
    return _banner_slide(deck, title)


def add_closing_slide(deck, text):
    """Closing slide (e.g. the thanks line): no page number."""
    return _banner_slide(deck, text)


def add_title_slide(deck, title=None, subtitle=None, lines=()):
    """Talk title, then affiliation/author/date lines. No page number."""
    slide = new_slide(deck, numbered=False)
    tf = _textbox(slide, 0.9, 2.1, 11.5, 1.6, anchor=MSO_ANCHOR.BOTTOM,
                  align=PP_ALIGN.CENTER).text_frame
    to_runs(tf.paragraphs[0], title, 40, bold=True)
    if subtitle:
        tf = _textbox(slide, 0.9, 3.8, 11.5, 0.6, align=PP_ALIGN.CENTER).text_frame
        to_runs(tf.paragraphs[0], subtitle, 24)
    lines = list(lines)
    if lines:
        tf = _textbox(slide, 0.9, 4.7, 11.5, 1.4, align=PP_ALIGN.CENTER).text_frame
        for i, line in enumerate(lines):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            para.alignment = PP_ALIGN.CENTER
            to_runs(para, line, 20)
    return slide


# --------------------------------------------------------------------------- #
# evidence
# --------------------------------------------------------------------------- #

def _aspect(src):
    """width/height of an image, or None when Pillow is missing or the file is odd."""
    if Image is None:
        return None
    try:
        if hasattr(src, "seek"):
            src.seek(0)
        with Image.open(src) as im:
            w, h = im.size
        return (w / h) if h else None
    finally:
        if hasattr(src, "seek"):
            src.seek(0)


def fit_images(slide, zone, images, gap=0.25, rows=1):
    """Lay images across the evidence zone, aspect-fit so nothing is stretched.

    Each entry is a dict with ``path`` (str/Path) or ``data`` (bytes/file-like),
    and an optional ``caption``.
    """
    images = list(images)
    if not images:
        return {"cell_w": 0.0, "cell_h": 0.0}
    cols = -(-len(images) // rows)  # ceil
    cap_h = 0.35 if any(im.get("caption") for im in images) else 0.0
    cell_w = (zone["w"] - gap * (cols - 1)) / cols
    cell_h = (zone["h"] - gap * (rows - 1)) / rows
    img_h = cell_h - cap_h

    for i, im in enumerate(images):
        r, c = divmod(i, cols)
        x = zone["x"] + c * (cell_w + gap)
        y = zone["y"] + r * (cell_h + gap)
        src = im.get("path")
        if src is None:
            data = im.get("data")
            if data is None:
                raise ValueError("each image needs a 'path' or 'data'")
            src = io.BytesIO(data) if isinstance(data, (bytes, bytearray)) else data
        else:
            src = str(Path(src))

        ar = _aspect(src)
        if ar:  # contain: fit inside the cell, then centre what is left over
            draw_w, draw_h = (cell_w, cell_w / ar) if cell_w / ar <= img_h else (img_h * ar, img_h)
        else:
            draw_w, draw_h = cell_w, img_h
        slide.shapes.add_picture(
            src,
            Inches(x + (cell_w - draw_w) / 2), Inches(y + (img_h - draw_h) / 2),
            Inches(draw_w), Inches(draw_h),
        )
        if im.get("caption"):
            tf = _textbox(slide, x, y + img_h, cell_w, cap_h,
                          align=PP_ALIGN.CENTER).text_frame
            to_runs(tf.paragraphs[0], im["caption"], G["caption"]["font_size"])
    return {"cell_w": cell_w, "cell_h": cell_h}


def red_box(slide, x, y, w, h, thickness=2):
    """Red enclosure linking a summary sentence to part of the evidence (rule 8)."""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y),
                                   Inches(w), Inches(h))
    shape.fill.background()
    shape.line.color.rgb = RGBColor.from_string(RED)
    shape.line.width = Pt(thickness)
    shape.shadow.inherit = False
    return shape


def balloon(slide, text, x, y, w, h=0.9, point_to=None):
    """Balloon for supplementary explanation kept out of the summary (rule 9)."""
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y),
                                   Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string("F2F2F2")
    shape.line.color.rgb = RGBColor.from_string(GRAY)
    shape.line.width = Pt(1)
    shape.shadow.inherit = False
    try:
        shape.adjustments[0] = max(0.0, min(0.5, 0.1 / min(w, h)))
    except (IndexError, ZeroDivisionError):
        pass

    tf = _textbox(slide, x + 0.1, y, w - 0.2, h, anchor=MSO_ANCHOR.MIDDLE,
                  align=PP_ALIGN.CENTER).text_frame
    to_runs(tf.paragraphs[0], text, G["caption"]["font_size"])

    if point_to:
        tx, ty = point_to["x"], point_to["y"]
        conn = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT,
            Inches(x + w / 2), Inches(y + h), Inches(tx), Inches(ty),
        )
        conn.line.color.rgb = RGBColor.from_string(GRAY)
        conn.line.width = Pt(1.5)
        conn.line._get_or_add_ln().append(parse_xml(
            '<a:tailEnd xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'type="triangle"/>'
        ))
    return shape


def caption(slide, text, x, y, w, h=0.35, align="left"):
    """14pt caption / source note. Never for content that belongs in the summary."""
    alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
                 "right": PP_ALIGN.RIGHT}.get(align, PP_ALIGN.LEFT)
    tf = _textbox(slide, x, y, w, h, align=alignment).text_frame
    to_runs(tf.paragraphs[0], text, G["caption"]["font_size"])
    return tf
