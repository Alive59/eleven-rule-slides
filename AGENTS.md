# AGENTS.md

Runtime-neutral entry point for this repository. Claude reads `SKILL.md` through
its skill loader; Codex and other agents that look for `AGENTS.md` land here.

**`SKILL.md` is the specification.** Read it before building or reviewing a deck —
it carries the workflow, the content-precedence rule, the fixed zone geometry, and
how the summary sentences must be written. This file only says which toolchain to
reach for and how to verify the result.

## Pick a generator

| You have | Use |
|---|---|
| Python + `python-pptx` | `scripts/slide_kit.py` |
| Node + `pptxgenjs` | `scripts/slide_kit.js` |
| Both | either; they produce the same geometry |

ChatGPT Code Interpreter, Codex sandboxes without npm, and CI images with no Node
all fall in the first row. `pip install -r requirements.txt` is the only setup.

## The loop

1. Read `SKILL.md`, then `references/eleven-rules.md`.
2. Ask which fonts the deck uses before generating anything (workflow step 0).
   Default pair: Calibri for Latin, Yu Gothic for Japanese.
3. Write the content — title, 2–4 summary sentences, the evidence for each — in
   plain text before touching either kit.
4. Generate with the kit. Never place text by hand at coordinates the kit already
   owns; that is how the geometry rules get broken.
5. Audit: `python scripts/check_rules.py deck.pptx --font-en Calibri --font-ja "Yu Gothic"`.
   Every FAIL must be fixed. Every WARN needs a stated reason.
6. Render and look at the slides if the sandbox can
   (`references/toolchain.md` has the commands). If it cannot, say so — do not
   report a visual check you did not perform.

## Hard constraints

- Do not edit content the user did not ask you to touch, even when it violates a
  rule. Report it and let them decide (SKILL.md, *Content precedence*).
- Title 32pt and up, body 20pt and up, page number 28pt pure black at the
  right-top. Cover, chapter-separator and closing slides carry no page number.
- Two font faces per deck, one Latin and one Japanese, written per run as
  `<a:latin>` and `<a:ea>`. Use the kit's `to_runs()` / `toRuns()`; a raw
  `run.font.name` styles only half of a mixed sentence.
- Red is reserved for enclosures that link a claim to its evidence (rule 8).
- Library footguns for both kits: `references/toolchain.md`.

## Repository map

```
SKILL.md                    the specification
AGENTS.md                   this file
references/eleven-rules.md  the eleven rules, with a do/don't per rule
references/toolchain.md     pptxgenjs and python-pptx footguns, render QA
scripts/slide_kit.py        generator (python-pptx)
scripts/slide_kit.js        generator (pptxgenjs)
scripts/check_rules.py      audit any .pptx against the mechanical rules
scripts/page_number.py      retrofit page numbers onto a deck you did not build
examples/example_deck.py    runnable example (Python)
examples/example_deck.js    runnable example (Node)
tests/test_python_kit.py    smoke tests: audit passes, geometry, font runs
gpt/instructions.md         condensed build for a ChatGPT Custom GPT
```
