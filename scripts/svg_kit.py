"""Vector visuals for eleven-rule decks: SVG built in code, placed as a vector picture.

See SKILL.md, *Visuals: vector first*. Style follows the building sketches of the
reference deck: flat gray faces, white edges, red only for defects.
"""
import os, subprocess
from lxml import etree
from pptx.util import Inches
from pptx.oxml.ns import qn

SVG_NS = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
SVG_EXT = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"
FONT = "'Yu Gothic','YuGothic','Noto Sans CJK JP','Hiragino Sans',sans-serif"

def _png_fallback(svg_path, px_w):
    png = os.path.splitext(svg_path)[0] + "_fb.png"
    try:
        import cairosvg
        cairosvg.svg2png(url=svg_path, write_to=png, output_width=px_w)
    except ImportError:
        subprocess.run(["rsvg-convert", "-w", str(px_w), "-o", png, svg_path], check=True)
    return png

def add_svg(slide, svg_path, x, y, w, h=None, name=None):
    """Place an SVG as a vector picture (svgBlip) with a PNG fallback, aspect kept.
    PowerPoint 2016+/365 draws the SVG; older viewers draw the PNG."""
    root = etree.parse(svg_path).getroot()
    vb = [float(v) for v in root.get("viewBox").split()]
    aspect = vb[3] / vb[2]
    if h is None or h / w > aspect:
        h = w * aspect
    else:
        w = h / aspect
    pic = slide.shapes.add_picture(_png_fallback(svg_path, int(w * 200)),
                                   Inches(x), Inches(y), Inches(w), Inches(h))
    svg_part_rid = _relate_svg(slide, svg_path)
    blip = pic._element.find(".//" + qn("a:blip"))
    ext_lst = etree.SubElement(blip, qn("a:extLst"))
    ext = etree.SubElement(ext_lst, qn("a:ext"), uri=SVG_EXT)
    sb = etree.SubElement(ext, "{%s}svgBlip" % SVG_NS, nsmap={"asvg": SVG_NS})
    sb.set(qn("r:embed"), svg_part_rid)
    if name:
        pic.name = name
    return pic

def _relate_svg(slide, svg_path):
    from pptx.opc.package import Part
    from pptx.opc.packuri import PackURI
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    pkg = slide.part.package
    partname = pkg.next_partname("/ppt/media/image%d.svg")
    with open(svg_path, "rb") as f:
        part = Part(PackURI(str(partname)), "image/svg+xml", pkg, f.read())
    return slide.part.relate_to(part, RT.IMAGE)

# --- building sketch (oblique view, flat faces, white edges) ------------------------
FACE = {"front": "#D9D9D9", "side": "#BFBFBF", "roof": "#8C8C8C", "roof_back": "#A6A6A6"}
RED, INK, GRAY = "#B2241E", "#222222", "#7F7F7F"

def _poly(pts, fill, stroke="#FFFFFF", sw=1.5, dash=None):
    d = f' stroke-dasharray="{dash}"' if dash else ""
    p = " ".join(f"{a:.1f},{b:.1f}" for a, b in pts)
    return f'<polygon points="{p}" fill="{fill}" stroke="{stroke}" stroke-width="{sw}"{d}/>'

def gable_building(x, y, w=165, wall=52, ridge=61, depth=40, slope=34, missing=None):
    """Gable-roof building in oblique view; front face bottom-left at (x, y), y = ground.
    missing=(i, n): roof split into n strips, strip i drawn as a red dashed gap with a
    red circle at its peak (the 'roof face missing' defect). Returns SVG markup."""
    top, ry, h = y - wall, y - wall - ridge, depth / 2
    back = ry + 0.8 * slope
    out = [
        _poly([(x + depth, back), (x + w + depth, back), (x + w + h, ry), (x + h, ry)], FACE["roof_back"]),
        _poly([(x + w, y), (x + w + depth, y - slope), (x + w + depth, back), (x + w + h, ry), (x + w, top)], FACE["side"]),
        _poly([(x, y), (x + w, y), (x + w, top), (x, top)], FACE["front"]),
    ]
    n = missing[1] if missing else 1
    sw_ = w / n
    for k in range(n):
        a, b = x + k * sw_, x + (k + 1) * sw_
        pts = [(a, top), (b, top), (b + h, ry), (a + h, ry)]
        gap = missing and k == missing[0]
        out.append(_poly(pts, "#FFFFFF", RED, 2.5, "7,5") if gap else _poly(pts, FACE["roof"]))
    if missing:
        cx = x + (missing[0] + 0.5) * sw_ + h
        out.append(f'<circle cx="{cx:.1f}" cy="{ry:.1f}" r="13" fill="none" stroke="{RED}" stroke-width="2.5"/>')
    return "\n".join(out)

def label(x, y, text, size=20, bold=False, color=INK, anchor="middle"):
    return (f'<text x="{x}" y="{y}" font-family="{FONT}" font-size="{size}" '
            f'font-weight="{"bold" if bold else "normal"}" fill="{color}" text-anchor="{anchor}">{text}</text>')

def svg_doc(width, height, body):
    """viewBox units: plan 100 units per inch of placed width, so font-size 20 = 14.4 pt."""
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{width*2}" height="{height*2}" '
            f'viewBox="0 0 {width} {height}">{body}</svg>')
