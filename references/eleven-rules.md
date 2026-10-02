# The eleven rules

Adapted from a lab guideline on presentation slides.

The premise: the presentation is not the research. The audience is not you and
cannot reconstruct the work from the slides. If a slide leaves them guessing, they
spend the Q&A on confirmation questions instead of real discussion. An individual
slide is not a free canvas — keep it simple and concrete, focused on the points and
the conclusion.

---

**1. Content composition.** Each slide consists of a title, summary sentences, and
supporting evidence (figures, tables). All three, every slide. A slide that is only
a figure forces the audience to derive the claim themselves; a slide that is only
text gives them nothing to check it against.

**2. Content placement.** Keep placement consistent across slides and follow the
audience's eye flow, top to bottom: title, then summary, then evidence. Mixed
layouts make the audience ask which part matters — the text or the figure.

**3. Content size.** Every text and figure large enough to be read by an old man at
the back of the room. Six small multiples where two would do is the usual failure:
the axis values and legends stop being recognizable. Cut the panels, enlarge the
survivors.

**4. Easy-to-understand title.** Express the title with keywords characteristic of
that slide. "5.2 Aggregated result [2/2]" tells the audience nothing;
"Result of extraction (Error pattern analysis)" tells them what they are about
to look at. Keep it compact: descriptive words only, as a noun phrase — never a
sentence. The finding goes in the summary sentences (rule 6), not in the title.

**5. Chapter separator slides.** Insert a slide between chapters ("Chapter 6: Case
study") so the individual titles can stay short and concrete instead of carrying
the chapter name as a prefix.

**6. Summary sentence composition.** Two to four sentences at most, stating the
conclusion clearly. Order: the big, important, positive thing first; then the
negative and exceptional things if they matter. Experiment conditions alone are not
a summary — the conclusion drawn from the evidence is what belongs there.

**7. Highlighted numerical values.** Summary sentences carry specific numbers,
highlighted, especially on problem-setting and results slides. "Some undesirable
cases need further improvement" is unmemorable; "regression analysis (R² = 0.86,
β = 0.73) on the aggregated counts" is what the audience repeats later.

**8. Linkage of summary and evidence.** When the relationship between a sentence
and the evidence is not obvious, highlight the relevant part of the evidence with a
red enclosure or a balloon. Without it the audience cannot tell which cell, bar, or
region the claim refers to.

**9. Separation of supplementary explanation.** Supplementary explanation does not
belong in the summary block. Put it somewhere visibly separate — a balloon attached
to the part it explains works well.

**10. Presentation style.** On stage, look at the audience and the screen, not the
laptop. Summary sentences exist so the talk can be natural rather than a script
read aloud — which also means they must be speakable, not dense paragraphs.

**11. Memorable talk in limited time.** Know the single most important conclusion,
with its number, that the talk must deliver. When time runs short, background and
that conclusion are what survive.

---

## Applying rules 1, 6, 7 together

The three content rules interact. A well-formed summary block for a results slide
looks like:

1. Condition sentence with the scale number: what was run, on how much data.
2. Headline result with the metric value.
3. Exception or failure mode with its share.

Anything else — why the metric was chosen, what the baseline is, how the split was
made — is supplementary (rule 9) and goes beside the figure, not in the block.

## Common violations to check for in a draft

- Title is a section number or a generic noun (rule 4).
- Title is a sentence — a verb, a subject, or ending punctuation — instead of a
  compact keyword phrase (rule 4).
- Five or more bullets, or one bullet that runs three lines (rule 6).
- Summary states what was done but never what came out of it (rules 1, 6).
- No digits anywhere in the summary of a results or problem-setting slide (rule 7).
- A table with a highlighted column but no sentence referring to it, or a sentence
  about "the proposed method" with nothing marked in the table (rule 8).
- Parenthetical definitions and caveats inside a summary bullet (rule 9).
- Figure shrunk to fit text that overran its zone (rules 2, 3).
