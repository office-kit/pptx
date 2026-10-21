# Talk mode: slides a speaker presents

A talk deck is a picture story the speaker talks over. The audience sees each slide for
seconds and must always know where in the story they are. It is not a document: what
the speaker says goes in the notes, and the slide carries only what must be seen.
Coverage, caveats and footnotes that a reader would need do not belong on it.

Most rules below come from minorun365's slide-story and writing-guide skills, which
were distilled from years of the author's own conference decks. See
[sources](sources.md). When the deck is Japanese, also apply [writing-ja.md](writing-ja.md);
where the two differ, this file wins for headlines.

## Before outlining

1. Confirm the official title, abstract and time slot from the event page. Everything the
   abstract promises must appear in the talk; check this at the outline and at delivery.
2. Ask the speaker what the audience should take home, who the audience is, and where the
   talk should surprise them. Collect the speaker's real material first: incidents,
   numbers, screenshots, logs, past decks on the same topic. A structure cannot generate
   the weight of a real story; never invent anecdotes, quotes or figures to fill a slot.
3. If the speaker has a past deck on the same subject, build from it: keep content, order
   and phrasing, and change only stale facts, minimal bridges and slide-level cuts.
   Present candidate cuts to the speaker as blocks of neighbouring slides, not as a list.
   When the user asks for the deck to be restructured, keep only its facts, numbers and
   history (the speaker notes are often the richest source) and rebuild order and wording.
4. Agree the picture-story outline (one line per slide) with the speaker before styling.

Do not pour content into a template's shape. Three "icon + bold heading + one line" cards
repeated on every slide looks finished and says nothing. Do not import document devices
(cards, comparison grids, hub diagrams) that put four thin blocks on one slide.

## Shape of the story

**Open from where the audience is**, not from a definition: voice their problem
(「最近 MCP ってよく聞きますよね！…分かりづらくないですか？」), a provocative quiz, a news hook,
or the speaker's own recent experience. After the self-introduction, one or two questions
to the audience work as the entrance to the body.

**Do not give the answer first.** Revealing information one step at a time keeps the
cognitive load low and the talk interesting.

- No agenda or "today I will talk about three things" slide, no upfront conclusion, no
  overview map at the start.
- When a slide asks a question, answer it on the following slides. Do not ask and answer
  on the same slide.
- Before adding a slide, check it is not a summary of what comes later.
- The talk still has one governing argument; keep it in the outline and the notes. The
  slides show one discovery at a time. If you want to write 「〜の 2 つです」 or
  「あとで証明します」 on a slide, it has become a declaration.

**Section dividers voice the audience's next question** instead of naming a chapter:
「え、じゃあ AI エージェントって何なの…？」「よし、書けた！…どこにデプロイする？」.

- The slide right after a divider starts answering it, and the answer is complete before
  the next divider. Delete dividers nothing answers.
- Only the question is on the divider: no number, sub-line or chapter label. At a large
  size it may wrap into two lines at a meaning boundary.
- Place one every 3–6 body slides; never more than two in a row. An opening question to
  the audience counts as a body slide.
- Never ask the same question twice to keep that rhythm. When one question genuinely
  covers a long stretch, a longer gap is better than a repeated divider.
- When the story spans a change (a migration, a redesign), say on each affected slide
  whether it shows before or after the change.
- With three or more chapters, show where the audience is with a small, quiet step bar
  at the top of the body slides (not on the divider and not as an extra slide). It
  shows position, never answers.

**Introduce a concept in three steps**: what it is (one-line definition, etymology if it
helps), what it makes better (before/after), then a rough one-liner. Explain abstract →
concrete → analogy. An analogy explains; it never decorates. After an analogy, derive the
next slide from it instead of switching to another framing.

**Body flow** that works for most technical talks: a concrete example first ("this is what
you can build") → the limit of the current way → the mechanism (one diagram) → how to do
it → where people get stuck → what solves it → the next step. Keep each claim next to its
example; do not put another topic between them.

**Staging**: use at least two or three of these per deck, from the speaker's real material.
A deck with none reads as a model student's.

- A diagram built up over 3–5 slides instead of shown complete
- A false ending and reversal
- A short dialogue that turns a protocol or negotiation into speech bubbles
- Anticipating a misconception (「もしかして、これを思い浮かべていませんか…？」)
- Objection and rebuttal in quick succession
- At least one honest limitation slide; never sell only
- The speaker's own name for an experience

**Close with an action**, not a summary slide: a call to action for the audience, then the
promotion or book cover if any. Do not put a "thank you" slide before a promotion slide.
If the material has no call to action, derive it from what the speaker said the audience
should take home, and ask the speaker to confirm the wording.

## One slide

| Item     | Rule                                                                                         |
| -------- | -------------------------------------------------------------------------------------------- |
| Amount   | One headline and 3–4 lines of body. Split beyond that.                                       |
| Headline | Tells what the slide is about the moment it appears. One line; rephrase instead of wrapping. |
| Body     | 3–4 items, one level, 1–2 sentences each.                                                    |
| Emphasis | At most one per slide; across the deck, about one per 2–3 body slides.                       |
| Figure   | One per slide, large. A figure-only slide may reuse the previous headline.                   |
| Table    | 2–3 simple columns.                                                                          |

Official copy the speaker must use verbatim (company or product introductions) is exempt
from the amount rule; keep it on as few slides as the copy allows.

When a figure is the star, text is the headline plus at most one intro line; say the
conclusion aloud. Do not add small text for what the screen already shows: captions
under logos or screenshots, grey afterthought lines, coloured closing lines, notes under
tables. The test: "Can the audience see it on screen already, or would it change their
judgement today?" A bottom band that restates the slide in other words is the same
mistake. One exception: if the deck will be published, a claim about someone else's
product may carry a single small source line, because the slide will be read without the
speaker.

## Headlines: mix the forms, then count

All-polite or all-plain headlines both look machine-made. Decide the form per slide from
its role before writing it; do not write neutral text and convert the ending afterwards.

| Role of the slide                        | Form                                 | Example                               |
| ---------------------------------------- | ------------------------------------ | ------------------------------------- |
| A figure, code or screenshot is the star | Short noun label (3–6 chars is fine) | こうなりがち／アーキテクチャ例        |
| A claim or finding                       | Plain-form statement                 | 令和の AI エージェントは 3 行で書ける |
| Surprise, good news, a turn              | Exclamation                          | デフォルトでストリーミング対応！      |
| The audience's inner voice               | Spoken monologue                     | よし、エージェント書けた！            |
| A rebuttal, promotion or close           | Polite form                          | …人生そんなに簡単じゃないんです       |
| A bridge to the next slide               | Cut off mid-thought                  | 技術的な下地は整ったが…               |
| Introducing a term                       | Descriptive phrase + 「term」        | AI に記憶をもたせる「メモリー」       |
| Divider question                         | 「〜の？」                           | 何を用意すればいいの？                |

Two kinds of endings sound wrong in Japanese:

- A polite-form sentence with only its ending swapped to plain form (割に合いません →
  割に合わない). Rebuild it as a noun ending, 「〜しよう」 or 「！」 instead.
- Contracted colloquial endings (〜てる, 〜ばいい, 〜んです) used to sound friendly.
  Use 「OK」, 「〜の？」 or 「〜しよう」.

Uniformity is invisible slide by slide and obvious in a list. After drafting, run
`deck-text.mjs --mode talk` ([review.md](review.md)) and compare with hand-made decks by
the same author (two decks of 57–58 slides; the percentages count every slide,
dividers included):

| Metric                | Hand-made | Machine-uniform            |
| --------------------- | --------- | -------------------------- |
| Headlines with 「！」 | 7–20%     | 1%                         |
| Headlines with 「？」 | 10–15%    | 7%                         |
| Polite-form headlines | 10–12%    | 1%                         |
| Shortest headline     | 3–6 chars | every headline 15–24 chars |

These are one speaker's numbers: treat a large gap as a signal to reread, not a quota.
Many question dividers push 「？」 above the range; that is fine when the questions differ
in shape.
Also look for the same pattern three slides in a row: comma-pause two-part headlines
(「覚える言葉は、この 3 つだけ」), headlines that start with the same connective
(しかも／実は／ちなみに), and colon headlines (「品質：」「量：」).

Slides are not README headings. On a slide, 「簡単」「〜するだけ！」 are legitimate: the
headline's job is to carry the audience to the next point, not to label a section.

## Quotes and speech bubbles

- Quote only from real logs. A plausible sentence presented as "what I actually typed" is
  fabrication; without a real source, rewrite the slide so it no longer claims one.
- Keep real quotes verbatim; do not polish them for looks.
- Do not put a clever aphorism in a coloured quote box. Boxes hold concrete steps, numbers,
  examples or real prompts.
- A speech bubble must not restate the slide's conclusion. It carries a concrete honest
  aside that the bullets do not.

## Notes and timing

Put the spoken explanation, sources and caveats in the slide notes (`Slide notes`), in
the speaker's voice. The speaker, not the slide, says the details.

Estimate the finish time when the outline settles, whenever the slide count moves by 5+,
and before delivery. Count body slides separately from dividers and the cover (about 8
seconds each).

| Setting                             | Seconds per body slide |
| ----------------------------------- | ---------------------- |
| 5-minute lightning talk             | ~30                    |
| ~15-minute slot, screenshot-heavy   | ~29                    |
| Developer conference, ~40 minutes   | ~37                    |
| Exhibition or non-engineer audience | ~46                    |

The unit times exclude demos and Q&A: add those as fixed blocks (+0.5 min per live demo
for switching). If you do not know whether Q&A is inside the slot, ask; if you cannot,
plan the talk for the slot minus 5 minutes and say so. Aim to finish about 5 minutes
early in a venue. Rather than cutting early, mark up to two checkpoints:
"if we pass this divider after N minutes, skip that slide later". Never skip a divider.
Replace these unit times with the speaker's own once measured.

## Final check for talk decks

1. No slide exceeds one headline plus 3–4 body lines.
2. No headline is a bare topic label unless a figure is the star.
3. No slide has two figures or two emphasised spots.
4. Every divider starts being answered on the next slide and is fully answered before the
   next divider; no question is asked twice.
5. No agenda, upfront summary or overview-as-answer slipped in.
6. The headline mix and pattern runs look like a person's, not a template's.
7. The abstract's promises all appear, the title's promise is paid off by the end, and the
   time estimate fits the slot.
8. Numbers, mechanisms and terms agree across pages ([review.md](review.md) step 4).
