"""Smoke tests for the python-pptx kit. Run with pytest, or directly:

    python tests/test_python_kit.py

Asserts what the port has to preserve: the audit passes on a generated deck, the
geometry matches slide_kit.js, mixed Latin/Japanese runs carry both typefaces, and
the page number lands on content slides only.
"""
import re
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from pptx import Presentation  # noqa: E402
from pptx.util import Inches  # noqa: E402

from page_number import add_page_number  # noqa: E402
from slide_kit import (  # noqa: E402
    add_chapter_slide, add_closing_slide, add_content_slide, add_title_slide,
    balloon, caption, create_deck, fit_images, red_box,
)

EMU_IN = 914400.0
SUMMARY = [
    "Trained on **1.3M** samples; IoU reached **0.79** on the held-out split.",
    "Failures concentrate in the **dense** class (**12%** of samples).",
    "対象は **4** 地域で、F1 は **0.82** に到達した。",
]


def _figure(path, size=(1200, 700)):
    from PIL import Image
    Image.new("RGB", size, (210, 215, 230)).save(path)
    return path


def _build(tmp):
    fig = _figure(tmp / "fig.png")
    deck = create_deck(font_en="Calibri", font_ja="Yu Gothic", title="t", author="a")
    add_title_slide(deck, title="Example deck title", subtitle="sub",
                    lines=["Presenter", "Affiliation"])
    add_chapter_slide(deck, "Chapter 1: Experiment")
    slide, zone = add_content_slide(deck, title="Accuracy by input class",
                                    summary=SUMMARY, notes="note")
    fit_images(slide, zone, [{"path": str(fig), "caption": "Held-out split"}])
    red_box(slide, x=zone["x"] + 0.4, y=zone["y"] + 0.6, w=2.0, h=1.2)
    balloon(slide, text="Supplementary note (rule 9).", x=9.0, y=5.9, w=3.2, h=0.8,
            point_to={"x": 7.5, "y": 5.0})
    caption(slide, text="Source: internal", x=0.6, y=3.0, w=2.4)
    add_closing_slide(deck, "Thank you")
    out = tmp / "deck.pptx"
    deck.save(out)
    return out


def test_audit_passes():
    """The whole point of the port: a generated deck must audit clean."""
    with tempfile.TemporaryDirectory() as d:
        out = _build(Path(d))
        r = subprocess.run(
            [sys.executable, str(ROOT / "scripts" / "check_rules.py"), str(out),
             "--font-en", "Calibri", "--font-ja", "Yu Gothic"],
            capture_output=True, text=True,
        )
        assert r.returncode == 0, r.stdout + r.stderr
        assert "0 FAIL, 0 WARN" in r.stdout, r.stdout


def test_geometry_and_page_numbers():
    with tempfile.TemporaryDirectory() as d:
        prs = Presentation(str(_build(Path(d))))
        assert round(prs.slide_width / EMU_IN, 2) == 13.33
        assert round(prs.slide_height / EMU_IN, 2) == 7.5

        numbered = [i for i, s in enumerate(prs.slides, 1)
                    if any("slidenum" in sh._element.xml for sh in s.shapes)]
        assert numbered == [3], f"only content slides carry a number, got {numbered}"

        content = prs.slides[2]
        boxes = {round(sh.top / EMU_IN, 2): round(sh.left / EMU_IN, 2)
                 for sh in content.shapes if sh.has_text_frame}
        assert boxes.get(0.28) == 0.6, "title zone"
        assert boxes.get(1.15) == 0.6, "summary zone"

        title = min((sh for sh in content.shapes if sh.has_text_frame),
                    key=lambda sh: sh.top if "slidenum" not in sh._element.xml else 10**9)
        assert all(r.font.size.pt >= 32 for p in title.text_frame.paragraphs
                   for r in p.runs)


def test_mixed_script_runs_carry_both_faces():
    with tempfile.TemporaryDirectory() as d:
        prs = Presentation(str(_build(Path(d))))
        runs = [r for sh in prs.slides[2].shapes if sh.has_text_frame
                for p in sh.text_frame.paragraphs for r in p.runs
                if "対象" in sh.text_frame.text]
        assert runs, "japanese summary line not found"
        for r in runs:
            xml = r._r.xml
            lat = re.search(r'<a:latin[^>]*typeface="([^"]*)"', xml)
            ea = re.search(r'<a:ea[^>]*typeface="([^"]*)"', xml)
            assert ea and ea.group(1) == "Yu Gothic", xml
            assert lat and lat.group(1) in ("Calibri", "Yu Gothic"), xml
        bolded = [r.text for r in runs if r.font.bold]
        assert "4" in bolded and "0.82" in bolded, bolded


def test_summary_cap_is_enforced():
    deck = create_deck()
    for bad, msg in ((dict(title="t", summary=["a"] * 5), "rule 6"),
                     (dict(title="", summary=["a"]), "rule 4"),
                     (dict(title="Accuracy improves with LoD2 input.", summary=["a"]),
                      "rule 4"),
                     (dict(title="LoD2入力で精度が向上した", summary=["a"]), "rule 4"),
                     (dict(title="We compare baselines", summary=["a"]), "rule 4")):
        try:
            add_content_slide(deck, **bad)
        except ValueError as e:
            assert msg in str(e), e
        else:
            raise AssertionError(f"expected ValueError for {bad}")


def test_retrofit_skips_covers():
    """page_number.add_page_number keeps its heuristic and its old signature."""
    with tempfile.TemporaryDirectory() as d:
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
        blank = prs.slide_layouts[6]
        cover = prs.slides.add_slide(blank)
        tb = cover.shapes.add_textbox(Inches(1), Inches(3), Inches(11), Inches(1.5))
        tb.text_frame.text = "Chapter 1"
        from pptx.util import Pt
        tb.text_frame.paragraphs[0].runs[0].font.size = Pt(40)

        body = prs.slides.add_slide(blank)
        tb2 = body.shapes.add_textbox(Inches(0.6), Inches(1.15), Inches(12), Inches(1))
        tb2.text_frame.text = "A summary sentence with 12 in it."
        tb2.text_frame.paragraphs[0].runs[0].font.size = Pt(20)

        assert add_page_number(cover, 1, 13.333) is False
        assert add_page_number(body, 2, 13.333) is True
        out = Path(d) / "retro.pptx"
        prs.save(out)
        prs2 = Presentation(str(out))
        assert not any("slidenum" in sh._element.xml for sh in prs2.slides[0].shapes)
        assert any("slidenum" in sh._element.xml for sh in prs2.slides[1].shapes)


if __name__ == "__main__":
    fails = 0
    for name, fn in sorted(globals().items()):
        if name.startswith("test_") and callable(fn):
            try:
                fn()
                print(f"PASS {name}")
            except AssertionError as e:
                fails += 1
                print(f"FAIL {name}: {e}")
    print(f"\n{fails} failure(s)")
    sys.exit(1 if fails else 0)


def test_title_sentence_reason():
    from slide_kit import title_sentence_reason
    for ok in ("Accuracy by input class", "Comparison with baselines in dense areas",
               "入力クラス別の再構成精度", "Dataset and study area", "Error patterns"):
        assert title_sentence_reason(ok) is None, ok
    for bad in ("Results are good.", "精度が向上した", "提案手法の評価です",
                "Our method outperforms baselines", "What is BMQI?"):
        assert title_sentence_reason(bad), bad
