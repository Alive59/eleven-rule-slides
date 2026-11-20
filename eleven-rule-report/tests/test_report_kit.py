"""Smoke test for report_kit: builds a report slide on top of slide_kit and checks
that the charts carry no title or legend, use the term-sheet colors, and that the
deck still passes the base audit. Run: python eleven-rule-report/tests/test_report_kit.py
"""
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parents[1]
ROOT = HERE.parent
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(HERE / "scripts"))

from pptx import Presentation  # noqa: E402
from slide_kit import (add_closing_slide, add_content_slide, add_title_slide,  # noqa: E402
                       create_deck)
from report_kit import (PALETTE, chart_violations, compare_bars, flat_violations,  # noqa: E402
                        flatten_shapes, grouped_bars, label_under, plain_fill,
                        report_table, slide_legend)
from slide_kit import balloon, red_box  # noqa: E402


def build(path):
    deck = create_deck(font_en="Calibri", font_ja="Yu Gothic")
    add_title_slide(deck, title="進捗報告", lines=["定例会"])
    slide, _ = add_content_slide(deck, title="評価対象全体の水密性", summary=[
        "水密性：生成モデルが閉じた面で構成される建築物の割合 (%)",
        "評価対象 **2,500 棟**、面数で 5 区分",
        "水密性 **78.2 → 84.5%**",
    ])
    compare_bars(slide, 0.6, 2.9, 3.8, 3.6, ["前回", "今回"], [78.2, 84.5], y_min=60)
    label_under(slide, 0.6, 3.8, "水密性 (%)")
    grouped_bars(slide, 4.8, 2.9, 7.9, 3.4, ["≤16", "17-32", "ALL"],
                 {"前回": [100, 99.3, 77.7], "今回": [100, 99.3, 84.1]})
    slide_legend(slide, 6.8, 6.5, ["前回", "今回"])
    balloon(slide, text="大型建築物", x=9.5, y=2.9, w=2.0, h=0.6, point_to={"x": 11.0, "y": 4.5})
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches
    block = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(1), Inches(1), Inches(0.5))
    block.shadow.inherit = True                    # a shape left with theme styling
    slide, _ = add_content_slide(deck, title="面数別の頂点誤差", summary=[
        "頂点誤差：対応頂点間距離 (m) の建築物平均",
        "水平誤差 **1.00 → 0.77 m**",
    ])
    report_table(slide, 0.6, 3.0, 12.0, 2.0, "面数", ["前回", "今回"],
                 ["水平誤差 (m)", "鉛直誤差 (m)"],
                 [("≤16", ["0.06", "0.70", "0.05", "0.71"]),
                  ("全体", ["1.00", "0.97", "0.77", "0.73"])],
                 bold={(0, 2), (1, 2), (1, 3)})
    red_box(slide, x=0.6, y=4.3, w=12.0, h=0.5)   # as if the user asked for it
    add_closing_slide(deck, "ご清聴ありがとうございました")
    assert flat_violations(deck.prs)  # the raw block fails
    flatten_shapes(deck.prs)
    deck.save(path)


def test_report_kit():
    with tempfile.TemporaryDirectory() as d:
        path = str(Path(d) / "report.pptx")
        build(path)
        prs = Presentation(path)
        charts = [sh.chart for s in prs.slides for sh in s.shapes if sh.has_chart]
        assert len(charts) == 2
        for ch in charts:
            assert not ch.has_title and not ch.has_legend
        xml = "".join(sh.chart.part.blob.decode("utf8") for s in prs.slides
                      for sh in s.shapes if sh.has_chart)
        for color in (PALETTE["前回"], PALETTE["今回"]):
            assert color in xml
        tables = [sh.table for s in prs.slides for sh in s.shapes if sh.has_table]
        assert len(tables) == 1 and tables[0].cell(0, 1).text == "前回"
        assert flat_violations(prs) == [], flat_violations(prs)
        assert chart_violations(prs) == [], chart_violations(prs)
        for ch in charts:                              # value labels keep the decimals
            assert 'formatCode="0.0"' in ch.part.blob.decode("utf8")
        r = subprocess.run([sys.executable, str(ROOT / "scripts" / "check_rules.py"), path],
                           capture_output=True, text=True)
        assert "| 0 FAIL" in r.stdout, r.stdout


if __name__ == "__main__":
    test_report_kit()
    print("ok")
