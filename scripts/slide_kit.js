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

// The page number lives once on the slide master (ppt/slideMasters/slideMaster1.xml)
// as an ordinary text box holding a slidenum field: every layout that shows master
// shapes inherits it, so every slide — including one added later in PowerPoint —
// prints its own number with no per-slide object. pptxgenjs can only emit
// per-slide number placeholders, so the master box is written into the package at
// export time. The ELEVEN_RULE_PLAIN layout hides master shapes (showMasterSp="0");
// only the title and acknowledgement slides use it.
const EMU = 914400;
const MASTER_NUM_NAME = 'Page Number (master)';

function masterNumberXml(id, font) {
  const p = G.pageNum;
  const sz = p.fontSize * 100;
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${MASTER_NUM_NAME}"/>` +
    '<p:cNvSpPr txBox="1"/><p:nvPr userDrawn="1"/></p:nvSpPr>' +
    `<p:spPr><a:xfrm><a:off x="${Math.round(p.x * EMU)}" y="${Math.round(p.y * EMU)}"/>` +
    `<a:ext cx="${Math.round(p.w * EMU)}" cy="${Math.round(p.h * EMU)}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>' +
    '<p:txBody><a:bodyPr wrap="none" rtlCol="0" anchor="t"><a:noAutofit/></a:bodyPr>' +
    '<a:lstStyle/><a:p><a:pPr algn="r"/>' +
    '<a:fld id="{1D0E7A29-9A6B-4C5E-9F0A-00000000A11D}" type="slidenum">' +
    `<a:rPr lang="en-US" sz="${sz}" b="0" dirty="0">` +
    '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>' +
    `<a:latin typeface="${font}"/><a:ea typeface="${font}"/><a:cs typeface="${font}"/>` +
    '</a:rPr><a:t>\u2039#\u203a</a:t></a:fld>' +
    `<a:endParaRPr lang="en-US" sz="${sz}"/></a:p></p:txBody></p:sp>`;
}

/** Rewrite an exported .pptx so the master carries the page number. */
async function moveNumberToMaster(data, font) {
  const JSZip = require('jszip'); // a pptxgenjs dependency, always installed with it
  const zip = await JSZip.loadAsync(data);
  const masterPath = 'ppt/slideMasters/slideMaster1.xml';
  let master = await zip.file(masterPath).async('string');
  if (!master.includes(MASTER_NUM_NAME)) {
    const ids = [...master.matchAll(/<p:cNvPr id="(\d+)"/g)].map((m) => +m[1]);
    const id = Math.max(1, ...ids) + 1;
    master = master.replace('</p:spTree>', masterNumberXml(id, font) + '</p:spTree>');
    zip.file(masterPath, master);
  }
  for (const path of Object.keys(zip.files)) {
    if (!/^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(path)) continue;
    let xml = await zip.file(path).async('string');
    if (!xml.includes('<p:cSld name="ELEVEN_RULE_PLAIN"')) continue;
    if (!/<p:sldLayout\b[^>]*showMasterSp=/.test(xml)) {
      xml = xml.replace(/<p:sldLayout\b/, '<p:sldLayout showMasterSp="0"');
      zip.file(path, xml);
    }
  }
  return zip;
}

/** New deck whose slide master carries the auto page number (top-right). */
function createDeck(opts = {}) {
  setFonts(opts);
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE';
  pres.defineSlideMaster({ title: 'ELEVEN_RULE', background: { color: 'FFFFFF' } });
  // Title and acknowledgement slides: same canvas, master shapes hidden, so no number.
  pres.defineSlideMaster({ title: 'ELEVEN_RULE_PLAIN', background: { color: 'FFFFFF' } });
  if (opts.author) pres.author = opts.author;
  if (opts.title) pres.title = opts.title;

  // Every export path (writeFile, write, stream) goes through exportPresentation.
  const font = EN_FONT;
  const exportRaw = pres.exportPresentation;
  pres.exportPresentation = async (props = {}) => {
    const raw = await exportRaw({ compression: props.compression, outputType: 'nodebuffer' });
    const zip = await moveNumberToMaster(raw, font);
    const compression = props.compression ? 'DEFLATE' : 'STORE';
    const type = props.outputType === 'STREAM' ? 'nodebuffer' : (props.outputType || 'blob');
    return zip.generateAsync({ type, compression });
  };
  return { pres, G };
}

// Rule 4: a title is a compact keyword phrase, never a sentence. The claim belongs
// in the summary block. Keep these patterns in sync with slide_kit.py and check_rules.py.
const TITLE_TERMINAL_RE = /[.。!！?？]\s*$/;
const TITLE_JA_PREDICATE_RE =
  /(です|ます|ました|ません|でした|である|だった|した|する|される|された|できる|できた|いる|ない|なった|なる)[。.!！?？]?\s*$/;
const TITLE_EN_VERB_RE = new RegExp(
  '\\b(is|are|was|were|has|have|had|does|do|did|can|could|will|would|should|' +
  'we|our|improves?|improved|outperforms?|outperformed|reduces?|reduced|' +
  'increases?|increased|decreases?|decreased|achieves?|achieved|shows?|showed|' +
  'enables?|enabled|yields?|yielded|reaches|reached|fails|failed|leads|led)\\b', 'i');

/** Why `title` reads as a sentence, or null if it is a keyword phrase. */
function titleSentenceReason(title) {
  const t = String(title || '').trim();
  if (TITLE_TERMINAL_RE.test(t)) return 'ends with sentence punctuation';
  if (TITLE_JA_PREDICATE_RE.test(t)) return 'ends with a Japanese predicate';
  const m = t.match(TITLE_EN_VERB_RE);
  if (m) return `contains the verb/subject '${m[0]}'`;
  return null;
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
  const why = titleSentenceReason(title);
  if (why) {
    throw new Error(`rule 4: title '${title}' reads as a sentence (${why}); ` +
      'use a compact keyword phrase and move the claim into the summary');
  }
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

/** Chapter separator (rule 5): title only, centered, numbered like any slide. */
function addChapterSlide(deck, title) {
  const slide = newSlide(deck);
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

/** Acknowledgement / closing slide ("ご清聴ありがとうございました" etc.): no page number. */
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
  fitImages, redBox, balloon, caption, toRuns, setFonts, titleSentenceReason,
  G, RED, GRAY,
  get EN_FONT() { return EN_FONT; },
  get JA_FONT() { return JA_FONT; },
};
