"""The 28pt black slide-number field at the right-top, kept on the slide master.

The number lives once on the slide master as an ordinary text box holding a
``slidenum`` field, so every slide -- including one added later in PowerPoint --
prints its own number without a per-slide object. The title slide and the
acknowledgement (closing) slide opt out by hiding master shapes
(``showMasterSp="0"``), on their layout or on the slide itself.

``install_master_page_number`` is what slide_kit.py uses. ``number_deck_on_master``
retrofits a finished deck. ``add_page_number`` is the older per-slide retrofit,
kept for callers that cannot touch the master.
"""
import copy
import re

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

MASTER_NUM_NAME = "Page Number (master)"

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
MASTER_SP = (
    '<p:sp xmlns:p="%s" xmlns:a="%s">'
    '<p:nvSpPr><p:cNvPr id="{id}" name="%s"/><p:cNvSpPr txBox="1"/>'
    '<p:nvPr userDrawn="1"/></p:nvSpPr>'
    '<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
    '<p:txBody><a:bodyPr wrap="none" rtlCol="0" anchor="t"><a:noAutofit/></a:bodyPr>'
    '<a:lstStyle/>'
    '<a:p><a:pPr algn="r"/>'
    '<a:fld id="{{1D0E7A29-9A6B-4C5E-9F0A-00000000A11D}}" type="slidenum">'
    '<a:rPr lang="en-US" sz="{sz}" b="0" dirty="0">'
    '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>'
    '<a:latin typeface="{font}"/><a:ea typeface="{font}"/><a:cs typeface="{font}"/>'
    '</a:rPr><a:t>\u2039#\u203a</a:t></a:fld>'
    '<a:endParaRPr lang="en-US" sz="{sz}"/></a:p></p:txBody></p:sp>'
) % (P, A, MASTER_NUM_NAME)

THANKS_RE = re.compile(
    r"thank|acknowledg|ご清聴|ありがとう|謝辞|感谢|谢谢|致谢", re.I)

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


# --------------------------------------------------------------------------- #
# master-level numbering
# --------------------------------------------------------------------------- #

def _sp_tree(part_element):
    return part_element.find(qn('p:cSld')).find(qn('p:spTree'))


def _next_shape_id(sp_tree):
    ids = [int(v) for v in sp_tree.xpath('.//p:cNvPr/@id') if v.isdigit()]
    return max(ids, default=1) + 1


def master_page_number(master):
    """The master's page-number text box (not a placeholder), or None."""
    for sp in _sp_tree(master._element).findall(qn('p:sp')):
        nv = sp.find(qn('p:nvSpPr'))
        is_ph = nv is not None and nv.find(qn('p:nvPr')).find(qn('p:ph')) is not None
        if not is_ph and 'type="slidenum"' in sp.xml:
            return sp
    return None


def install_master_page_number(master, slide_w, font='Calibri'):
    """Put the 28pt black slide-number field on ``master``. Idempotent.

    A placeholder on the master prints nothing on the slides, so this is an
    ordinary text box: every layout that shows master shapes inherits it, and the
    field resolves to each slide's own number.
    """
    existing = master_page_number(master)
    if existing is not None:
        return existing
    tree = _sp_tree(master._element)
    sp = parse_xml(MASTER_SP.format(
        id=_next_shape_id(tree),
        x=int(Inches(slide_w - BOX_W - RIGHT_PAD)), y=int(Inches(TOP)),
        cx=int(Inches(BOX_W)), cy=int(Inches(BOX_H)), sz=PT * 100, font=font))
    ext = tree.find(qn('p:extLst'))
    if ext is not None:
        ext.addprevious(sp)
    else:
        tree.append(sp)
    return sp


def hide_master_shapes(part):
    """Opt a layout or a slide out of the master's shapes (and so the number)."""
    part._element.set('showMasterSp', '0')


def plain_layout(prs, base_layout, name='Eleven Rule Plain'):
    """A copy of ``base_layout`` that hides master shapes, for the title and
    acknowledgement slides. Reused if the deck already has one by that name."""
    master = base_layout.slide_master
    for layout in master.slide_layouts:
        if layout.name == name:
            return layout
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    from pptx.opc.packuri import PackURI
    from pptx.parts.slide import SlideLayoutPart

    package = prs.part.package
    partname = package.next_partname('/ppt/slideLayouts/slideLayout%d.xml')
    element = copy.deepcopy(base_layout._element)
    element.find(qn('p:cSld')).set('name', name)
    element.set('showMasterSp', '0')
    part = SlideLayoutPart(PackURI(str(partname)), base_layout.part.content_type,
                           package, element)
    part.relate_to(master.part, RT.SLIDE_MASTER)
    rId = master.part.relate_to(part, RT.SLIDE_LAYOUT)
    id_lst = master._element.get_or_add_sldLayoutIdLst()
    used = [int(v) for v in prs.part._element.xpath('//p:sldMasterId/@id')]
    used += [int(v) for m in prs.slide_masters
             for v in m._element.xpath('.//p:sldLayoutId/@id')]
    new_id = max(used + [2147483647]) + 1
    entry = parse_xml('<p:sldLayoutId xmlns:p="%s" xmlns:r="%s" id="%d" r:id="%s"/>'
                      % (P, 'http://schemas.openxmlformats.org/officeDocument/2006/'
                         'relationships', new_id, rId))
    id_lst.append(entry)
    return master.slide_layouts[len(master.slide_layouts) - 1]


def is_title_or_thanks(slide, index, n_slides):
    """Title slide (the first) or acknowledgement slide (a cover-like slide that
    says thanks, or the last slide when it is cover-like)."""
    if index == 0:
        return True
    if not is_cover(slide):
        return False
    text = " ".join(sh.text_frame.text for sh in slide.shapes if sh.has_text_frame)
    return bool(THANKS_RE.search(text)) or index == n_slides - 1


def number_deck_on_master(prs, font='Calibri'):
    """Retrofit a finished deck: number on every master, off on title/thanks slides.

    Returns a report dict. Hiding master shapes on a slide also hides any other
    master graphics there (a logo, a footer band); those slides are listed under
    ``hidden_also`` so the caller can tell the user instead of hiding silently.
    Per-slide number fields already on the deck are left alone and listed under
    ``per_slide`` -- removing them is the user's call (content precedence).
    """
    slide_w = prs.slide_width / EMU
    for master in prs.slide_masters:
        install_master_page_number(master, slide_w, font)
    report = {"numbered": [], "exempt": [], "hidden_also": [], "per_slide": []}
    slides = list(prs.slides)
    for i, slide in enumerate(slides):
        if any('type="slidenum"' in sh._element.xml for sh in slide.shapes):
            report["per_slide"].append(i + 1)
        if is_title_or_thanks(slide, i, len(slides)):
            hide_master_shapes(slide)
            report["exempt"].append(i + 1)
            master = slide.slide_layout.slide_master
            others = [sp for sp in _sp_tree(master._element).findall(qn('p:sp'))
                      if sp.find(qn('p:nvSpPr')).find(qn('p:nvPr')).find(qn('p:ph')) is None
                      and 'type="slidenum"' not in sp.xml]
            others += _sp_tree(master._element).findall(qn('p:pic'))
            if others:
                report["hidden_also"].append(i + 1)
        else:
            report["numbered"].append(i + 1)
    return report
