/**
 * Example deck. Run from the repository root:
 *
 *   node examples/example_deck.js
 *   python scripts/check_rules.py example.pptx
 *
 * Shows the four slide kinds and how a red enclosure and a balloon attach a
 * summary sentence to the part of the figure it refers to.
 */
const path = require('path');
const {
  createDeck, addTitleSlide, addChapterSlide, addContentSlide, addClosingSlide,
  fitImages, redBox, balloon,
} = require(path.join(__dirname, '..', 'scripts', 'slide_kit.js'));

// Ask the user which fonts they want before hard-coding a pair here.
const deck = createDeck({ fontEn: 'Calibri', fontJa: 'Yu Gothic' });

addTitleSlide(deck, {
  title: 'Example deck title',
  lines: ['Presenter name', 'Affiliation', '2026-01-01'],
});

addChapterSlide(deck, 'Chapter 1: Experiment');

const { slide, zone } = addContentSlide(deck, {
  title: 'Accuracy by input class',
  summary: [
    'Trained on **1.3M** samples; IoU reached **0.79** on the held-out split.',
    'Failures concentrate in the **dense** class (**12%** of samples).',
    '対象は **4** 地域で、F1 は **0.82** に到達した。',
  ],
  notes: 'Lead with the headline number, then the exception.',
});

// Replace with real figure paths; fitImages sizes them to fit the evidence zone.
const figures = process.argv.slice(2).map((p) => ({ path: p }));
if (figures.length) {
  fitImages(slide, zone, figures);
  redBox(slide, { x: zone.x + 0.4, y: zone.y + 0.6, w: 2.0, h: 1.2 });
  balloon(slide, {
    text: 'Supplementary note, kept out of the summary (rule 9).',
    x: 9.0, y: 5.9, w: 3.2, h: 0.8,
    pointTo: { x: 7.5, y: 5.0 },
  });
}

addClosingSlide(deck, 'Thank you');

deck.pres.writeFile({ fileName: 'example.pptx' })
  .then(() => console.log('wrote example.pptx'));
