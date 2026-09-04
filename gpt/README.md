# Running this skill on GPT

Two ways, depending on whether you want a reusable GPT or a one-off chat.

## A. Custom GPT

1. **Instructions** — paste `gpt/instructions.md` into the *Instructions* box. It is
   written to fit the box on its own; it is a condensed build of `SKILL.md`, not a
   pointer to it.
2. **Knowledge** — upload, so the model can read the full text when it needs it:
   - `SKILL.md`
   - `references/eleven-rules.md`
   - `references/toolchain.md`
   - `scripts/slide_kit.py`
   - `scripts/page_number.py`
   - `scripts/check_rules.py`
3. **Capabilities** — enable **Code Interpreter & Data Analysis**. It is what runs
   the kit and the audit. Web browsing and DALL·E are not used.
4. In the chat, the GPT copies `slide_kit.py` and `page_number.py` out of Knowledge
   into its working directory (they must sit side by side — `slide_kit.py` imports
   `page_number.py`), then generates and audits there.

## B. Plain ChatGPT with Code Interpreter

Upload `slide_kit.py`, `page_number.py` and `check_rules.py` into the chat, keeping
them in the same directory, then paste `gpt/instructions.md` as your first message.

## Why the Python kit

Code Interpreter has `python-pptx` and no Node, no npm, and no network. It also has
no LibreOffice Impress filter, so a deck cannot be rendered to images there: the
audit plus a read-back of shape geometry is the whole verification budget, and a
GPT should say so rather than implying it looked at the slides.

`Pillow` is present, so `fit_images()` gets aspect-correct figures.
