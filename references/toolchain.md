# Toolchain notes

The footguns each generator library hides, and the render-QA loop. Read this when
your runtime has no general pptx skill of its own; where one exists (Claude ships
`/mnt/skills/public/pptx/SKILL.md`), that file is fuller and this is the short form.

## pptxgenjs (`scripts/slide_kit.js`)

- Hex colors carry **no** `#`: `'C00000'`, not `'#C00000'`. A `#` is silently kept
  and the color comes out wrong.
- Set `pres.layout` **before** adding any slide. Afterwards it is ignored and the
  deck stays 10 x 7.5 in.
- `isTextBox: true` on every `addText` you position yourself. Without it the box is
  a placeholder and PowerPoint may autofit the text below your minimum size.
- `shrinkText: false` on the title, for the same reason.
- Bullets belong on the **first run** of a paragraph and the line break on the
  **last**. Setting `bullet: true` on every run makes one paragraph per run, which
  the audit then counts as five summary lines instead of two.
- Charts validate their own data: every series needs `name`, `labels`, `values` of
  equal length, or `writeFile` rejects.
- `writeFile` returns a promise. Ending the script without awaiting it writes
  nothing.

## python-pptx (`scripts/slide_kit.py`)

- There is no API for defining a slide master or layout. The kit writes a real
  `slidenum` field onto each numbered slide instead; PowerPoint renumbers it on
  open, so reordering slides is still safe.
- `run.font.name` writes only `<a:latin>`. Japanese glyphs resolve through
  `<a:ea>`, so a mixed run styled with `font.name` alone renders the kana in the
  Latin face. `slide_kit.py` writes `<a:latin>`, `<a:ea>` and `<a:cs>` per run;
  do not bypass `to_runs()`.
- Child order inside `<a:rPr>` is fixed: fills, then `latin`, `ea`, `cs`. Appending
  in the wrong order gives a file PowerPoint refuses to open.
- `text_frame.auto_size = MSO_AUTO_SIZE.NONE` or PowerPoint may shrink text below
  the 20pt floor.
- `add_picture` with both `width` and `height` **stretches** the image. Compute the
  aspect ratio first (Pillow) and pass a fitted box — `fit_images()` does this.
- A slide added from a layout inherits that layout's placeholders, prompt text and
  all. `new_slide()` strips them.
- `Presentation()` starts at 10 x 7.5 in; set `slide_width`/`slide_height` to
  13.333 x 7.5 before adding slides.

## Fonts in a sandbox

Calibri and Yu Gothic are usually absent, so a render substitutes something else
and may drop bullet glyphs. Judge the preview on layout, overflow and figure size;
the audit reads the XML and is the authority on fonts and sizes.

## Render QA

```sh
soffice --headless --convert-to pdf --outdir out deck.pptx
pdftoppm -r 110 -png out/deck.pdf out/slide
```

Then look at every page. Overflow and figure placement are the two things
`check_rules.py` cannot see. Where LibreOffice is unavailable — ChatGPT Code
Interpreter has no Impress filter — the audit plus a read-back of the shape
geometry is what you have; say so rather than implying the deck was eyeballed.
