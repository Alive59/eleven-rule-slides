"""Example deck, python-pptx edition. Run from the repository root:

    python examples/example_deck.py [figure.png ...]
    python scripts/check_rules.py example_py.pptx

Mirrors examples/example_deck.js slide for slide: the four slide kinds, plus a red
enclosure and a balloon attaching a summary sentence to the part of the figure it
refers to.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from slide_kit import (  # noqa: E402
    add_chapter_slide, add_closing_slide, add_content_slide, add_title_slide,
    balloon, create_deck, fit_images, red_box,
)

# Ask the user which fonts they want before hard-coding a pair here.
deck = create_deck(font_en="Calibri", font_ja="Yu Gothic")

add_title_slide(
    deck,
    title="Example deck title",
    lines=["Presenter name", "Affiliation", "2026-01-01"],
)

add_chapter_slide(deck, "Chapter 1: Experiment")

slide, zone = add_content_slide(
    deck,
    title="Accuracy by input class",
    summary=[
        "Trained on **1.3M** samples; IoU reached **0.79** on the held-out split.",
        "Failures concentrate in the **dense** class (**12%** of samples).",
        "対象は **4** 地域で、F1 は **0.82** に到達した。",
    ],
    notes="Lead with the headline number, then the exception.",
)

# Replace with real figure paths; fit_images sizes them to fit the evidence zone.
figures = [{"path": p} for p in sys.argv[1:]]
if figures:
    fit_images(slide, zone, figures)
    red_box(slide, x=zone["x"] + 0.4, y=zone["y"] + 0.6, w=2.0, h=1.2)
    balloon(
        slide,
        text="Supplementary note, kept out of the summary (rule 9).",
        x=9.0, y=5.9, w=3.2, h=0.8,
        point_to={"x": 7.5, "y": 5.0},
    )

add_closing_slide(deck, "Thank you")

print("wrote", deck.save("example_py.pptx"))
