/**
 * slide_kit.js — geometry, typography and page numbering for eleven-rule decks.
 *
 * Everything the rules make non-negotiable (zone positions, font faces, minimum
 * sizes, top-right page number) lives here so a generator script cannot break it
 * by accident. Layout is LAYOUT_WIDE: 13.333 x 7.5 in.
 *
 *   const { createDeck, addContentSlide, fitImages } = require('./slide_kit.js');
 */

const pptxgen = require('pptxgenjs');

// Defaults; override per deck with createDeck({ fontEn, fontJa }) once the user
// has told you which fonts they want.
let EN_FONT = 'Calibri';
let JA_FONT = 'Yu Gothic';

function setFonts({ fontEn, fontJa } = {}) {
  if (fontEn) EN_FONT = fontEn;
  if (fontJa) JA_FONT = fontJa;
  return { fontEn: EN_FONT, fontJa: JA_FONT };
}
const RED = 'C00000';
const GRAY = '808080';
const BLACK = '000000';

// Japanese kana, CJK ideographs, CJK punctuation, fullwidth forms.
const CJK_RE = /[\u3000-\u303F\u3040-\u309F\u30A0-\u30FF\u3400-\u4DBF\u4E00-\u9FFF\uFF00-\uFFEF]/;

const G = {
  W: 13.333,
  H: 7.5,
  title: { x: 0.6, y: 0.28, w: 11.5, h: 0.75, fontSize: 32 },
  summary: { x: 0.6, y: 1.15, w: 12.1, h: 1.5, fontSize: 20 },
  evidence: { x: 0.6, y: 2.75, w: 12.1, h: 4.35 },
  caption: { fontSize: 14 },
  pageNum: { x: 11.98, y: 0.18, w: 0.95, h: 0.55, fontSize: 28 },
};

/** Split a string into runs so CJK spans get Yu Gothic and the rest Calibri.
 *  `**bold**` spans are emitted bold — use it for highlighted numerals (rule 7). */
function toRuns(text, opts = {}) {
  const runs = [];
  const parts = String(text).split(/(\*\*[^*]+\*\*)/g).filter(Boolean);
  for (const part of parts) {
    const bold = part.startsWith('**') && part.endsWith('**');
    const body = bold ? part.slice(2, -2) : part;
    let buf = '';
    let bufCjk = null;
    const flush = () => {
      if (!buf) return;
      runs.push({
        text: buf,
        options: Object.assign({}, opts, {
          fontFace: bufCjk ? JA_FONT : EN_FONT,
          bold: bold || opts.bold === true,
        }),
      });
      buf = '';
    };
    for (const ch of body) {
      const isCjk = CJK_RE.test(ch);
      if (bufCjk === null) bufCjk = isCjk;
      if (isCjk !== bufCjk) {
        flush();
        bufCjk = isCjk;
      }
      buf += ch;
    }
    flush();
  }
  return runs.length ? runs : [{ text: ' ', options: Object.assign({}, opts, { fontFace: EN_FONT }) }];
}

/** New deck with the master that carries the auto page number (top-right). */
function createDeck(opts = {}) {
  setFonts(opts);
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE';
  pres.defineSlideMaster({
    title: 'ELEVEN_RULE',
    background: { color: 'FFFFFF' },
    slideNumber: {
      x: G.pageNum.x,
      y: G.pageNum.y,
      w: G.pageNum.w,
      h: G.pageNum.h,
      align: 'right',
      fontFace: EN_FONT,
      fontSize: G.pageNum.fontSize,
      color: BLACK,
    },
  });
  // Cover, chapter separator and closing slides get the same canvas without the
  // number — the audience never needs to cite those pages.
  pres.defineSlideMaster({
    title: 'ELEVEN_RULE_PLAIN',
    background: { color: 'FFFFFF' },
  });
  if (opts.author) pres.author = opts.author;
  if (opts.title) pres.title = opts.title;
  return { pres, G };
}

function newSlide(deck, numbered = true) {
  return deck.pres.addSlide({ masterName: numbered ? 'ELEVEN_RULE' : 'ELEVEN_RULE_PLAIN' });
}

/**
 * Content slide: title + 2-4 summary sentences + reserved evidence zone.
 * Returns { slide, zone } — put figures in `zone` with fitImages().
 */
function addContentSlide(deck, { title, summary = [], notes } = {}) {
  if (!title) throw new Error('rule 4: every slide needs a keyword title');
  if (summary.length > 4) {
    throw new Error(`rule 6: ${summary.length} summary sentences; split the slide (max 4)`);
  }
  const slide = newSlide(deck);

  slide.addText(toRuns(title, { fontSize: G.title.fontSize, bold: true, color: BLACK }), {
    x: G.title.x, y: G.title.y, w: G.title.w, h: G.title.h,
    valign: 'middle', margin: 0, isTextBox: true, shrinkText: false,
  });

  if (summary.length) {
    const runs = [];
    summary.forEach((line, i) => {
      const items = toRuns(line, { fontSize: G.summary.fontSize, color: BLACK });
      // pptxgenjs makes one paragraph per bullet: the bullet flag goes on the
      // first run only, the line break on the last. Setting bullet on every run
      // splits each sentence into one paragraph per run.
      items.forEach((r, j) => {
        if (j === 0) {
          r.options.bullet = true;
          r.options.paraSpaceAfter = 6;
        }
        if (j === items.length - 1 && i !== summary.length - 1) r.options.breakLine = true;
        runs.push(r);
      });
    });
    slide.addText(runs, {
      x: G.summary.x, y: G.summary.y, w: G.summary.w, h: G.summary.h,
      valign: 'top', margin: 0, isTextBox: true,
    });
  }

  if (notes) slide.addNotes(notes);
  return { slide, zone: Object.assign({}, G.evidence) };
}

/** Chapter separator (rule 5): title only, centered, no page number. */
function addChapterSlide(deck, title) {
  const slide = newSlide(deck, false);
  slide.addText(toRuns(title, { fontSize: 40, bold: true, color: BLACK }), {
    x: 1.0, y: 2.8, w: 11.3, h: 1.9,
    align: 'center', valign: 'middle', margin: 0, isTextBox: true,
  });
  return slide;
}

/** Title slide: talk title, then affiliation/author/date lines. */
function addTitleSlide(deck, { title, subtitle, lines = [] } = {}) {
  const slide = newSlide(deck, false);
  slide.addText(toRuns(title, { fontSize: 40, bold: true, color: BLACK }), {
    x: 0.9, y: 2.1, w: 11.5, h: 1.6, align: 'center', valign: 'bottom', margin: 0, isTextBox: true,
  });
  if (subtitle) {
    slide.addText(toRuns(subtitle, { fontSize: 24, color: BLACK }), {
      x: 0.9, y: 3.8, w: 11.5, h: 0.6, align: 'center', margin: 0, isTextBox: true,
    });
  }
  if (lines.length) {
    slide.addText(toRuns(lines.join('\n'), { fontSize: 20, color: BLACK }), {
      x: 0.9, y: 4.7, w: 11.5, h: 1.4, align: 'center', margin: 0, isTextBox: true,
    });
  }
  return slide;
}

/** Closing slide ("ご清聴ありがとうございました" etc.): no page number. */
function addClosingSlide(deck, text) {
  const slide = newSlide(deck, false);
  slide.addText(toRuns(text, { fontSize: 40, bold: true, color: BLACK }), {
    x: 1.0, y: 2.8, w: 11.3, h: 1.9,
    align: 'center', valign: 'middle', margin: 0, isTextBox: true,
  });
  return slide;
}

/**
 * Lay images across the evidence zone in one row (or `rows` rows), `contain`-sized
 * so nothing is stretched. Each entry: { path | data, caption }.
 */
function fitImages(slide, zone, images, opts = {}) {
  const gap = opts.gap != null ? opts.gap : 0.25;
  const rows = opts.rows || 1;
  const cols = Math.ceil(images.length / rows);
  const capH = images.some((im) => im.caption) ? 0.35 : 0;
  const cellW = (zone.w - gap * (cols - 1)) / cols;
  const cellH = (zone.h - gap * (rows - 1)) / rows;
  images.forEach((im, i) => {
    const r = Math.floor(i / cols);
    const c = i % cols;
    const x = zone.x + c * (cellW + gap);
    const y = zone.y + r * (cellH + gap);
    const imgH = cellH - capH;
    const img = { x, y, w: cellW, h: imgH, sizing: { type: 'contain', w: cellW, h: imgH } };
    if (im.path) img.path = im.path;
    if (im.data) img.data = im.data;
    slide.addImage(img);
    if (im.caption) {
      slide.addText(toRuns(im.caption, { fontSize: G.caption.fontSize, color: BLACK }), {
        x, y: y + imgH, w: cellW, h: capH, align: 'center', valign: 'top', margin: 0, isTextBox: true,
      });
    }
  });
  return { cellW, cellH };
}

/** Red enclosure linking a summary sentence to part of the evidence (rule 8). */
function redBox(slide, { x, y, w, h, thickness = 2 }) {
  slide.addShape('rect', {
    x, y, w, h,
    fill: { type: 'none' },
    line: { color: RED, width: thickness },
  });
}

/** Balloon for supplementary explanation kept out of the summary (rule 9). */
function balloon(slide, { text, x, y, w, h = 0.9, pointTo }) {
  slide.addShape('roundRect', {
    x, y, w, h,
    fill: { color: 'F2F2F2' },
    line: { color: GRAY, width: 1 },
    rectRadius: 0.1,
  });
  slide.addText(toRuns(text, { fontSize: G.caption.fontSize, color: BLACK }), {
    x: x + 0.1, y, w: w - 0.2, h, valign: 'middle', align: 'center', margin: 0, isTextBox: true,
  });
  if (pointTo) {
    slide.addShape('line', {
      x: Math.min(x + w / 2, pointTo.x),
      y: Math.min(y + h, pointTo.y),
      w: Math.abs(pointTo.x - (x + w / 2)),
      h: Math.abs(pointTo.y - (y + h)),
      line: { color: GRAY, width: 1.5, endArrowType: 'triangle' },
      flipH: pointTo.x < x + w / 2,
      flipV: pointTo.y < y + h,
    });
  }
}

/** 14pt caption / source note. Never use this for content that belongs in the summary. */
function caption(slide, { text, x, y, w, h = 0.35, align = 'left' }) {
  slide.addText(toRuns(text, { fontSize: G.caption.fontSize, color: BLACK }), {
    x, y, w, h, align, valign: 'top', margin: 0, isTextBox: true,
  });
}

module.exports = {
  createDeck, newSlide, addContentSlide, addChapterSlide, addTitleSlide, addClosingSlide,
  fitImages, redBox, balloon, caption, toRuns, setFonts,
  G, RED, GRAY,
  get EN_FONT() { return EN_FONT; },
  get JA_FONT() { return JA_FONT; },
};
