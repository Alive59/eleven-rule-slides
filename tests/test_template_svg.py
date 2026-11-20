"""Smoke tests for use_template() and svg_kit. Run with pytest, or directly:

    python tests/test_template_svg.py
"""
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from pptx import Presentation  # noqa: E402
from pptx.oxml import parse_xml  # noqa: E402
from pptx.util import Inches  # noqa: E402

from slide_kit import add_content_slide, add_title_slide, create_deck, use_template  # noqa: E402
from svg_kit import add_svg, gable_building, label, svg_doc  # noqa: E402

BAND = ('<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:nvSpPr>'
        '<p:cNvPr id="99" name="band"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="0" y="6492240"/><a:ext cx="12192000" cy="365760"/>'
        '</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill>'
        '<a:srgbClr val="1F4E79"/></a:solidFill></p:spPr></p:sp>')


def _template(path):
    p = Presentation()
    p.slide_width, p.slide_height = Inches(13.333), Inches(7.5)
    p.slide_layouts.get_by_name("Title Only").shapes._spTree.append(parse_xml(BAND))
    p.slides.add_slide(p.slide_layouts[0])  # a sample slide to be dropped
    p.save(path)


def test_template_and_svg():
    with tempfile.TemporaryDirectory() as d:
        d = Path(d)
        _template(d / "lab.pptx")
        deck = create_deck(template=d / "lab.pptx")
        use_template(deck, "Title Only", "Title Slide")
        add_title_slide(deck, title="定例会")
        slide, _ = add_content_slide(deck, title="屋根面の欠落と補完",
                                     summary=["従来手法：屋根面の **1/3** が欠落", "生成AI：形状知識で補完"])
        svg = d / "sketch.svg"
        body = (gable_building(150, 330, missing=(1, 3)) + label(250, 380, "AIなし", bold=True)
                + gable_building(750, 330) + label(850, 380, "AIあり", bold=True))
        svg.write_text(svg_doc(1210, 435, body), encoding="utf-8")
        add_svg(slide, str(svg), 0.6, 2.75, 12.1, 4.35)
        out = d / "deck.pptx"
        deck.save(out)

        prs = Presentation(out)
        assert len(prs.slides) == 2, "template sample slide must be dropped"
        assert prs.slides[1].slide_layout.name == "Title Only"
        assert prs.slides[0].slide_layout._element.get("showMasterSp") == "0"
        names = zipfile.ZipFile(out).namelist()
        assert any(n.endswith(".svg") for n in names), "SVG part missing"
        assert "svgBlip" in prs.slides[1].shapes[-1]._element.xml
        r = subprocess.run([sys.executable, str(ROOT / "scripts" / "check_rules.py"), str(out)],
                           capture_output=True, text=True)
        assert " 0 FAIL" in r.stdout, r.stdout


if __name__ == "__main__":
    test_template_and_svg()
    print("ok")
