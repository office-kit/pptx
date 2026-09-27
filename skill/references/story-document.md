# Document mode: decks people read or decide from

A document deck is read without its author, or used in a meeting to reach a decision:
reports, proposals, decision memos, board packs, handouts. The reader wants the answer
first and must be able to follow the argument from the headlines alone.

The structure rules come from Barbara Minto's pyramid principle as operationalised by
tyroneross/pyramid-principle; the headline, bullet and table conventions come from
carnot-tech/consulting-pptx-skill (rules distilled from real consulting review
comments); message-line and density rules come from coji/natural-japanese. See
[sources](sources.md). When the deck is Japanese, also apply [writing-ja.md](writing-ja.md).

## Before outlining

1. Establish who reads it, in which meeting, and what they must understand or decide. If
   that is unknown, ask. If you cannot ask, state the reader you assumed when you deliver.
   When even the purpose is unclear, offer a one-page summary of the issues first: every
   extra slide adds exposure to errors and filler.
2. State the governing thought in one sentence: the decision, recommendation or finding.
   If you cannot, the material is insufficient; ask for it rather than writing around it.
3. Collect the facts the argument needs. Use only what the user supplied or you verified.
   A missing number or name is asked for, or shown as a visible `[要確認: …]`
   placeholder. Never invent a figure, customer, quote or source.
4. Write the storyline: one line per slide, in order. Lines include premises, facts and
   questions as well as answers; read in sequence they form one speech. Agree it with the
   user before building a deck of more than about ten slides.
5. Derive the slide count from the storyline. Each slide makes one claim, so count the
   claims each group of support needs (a mechanism a reader would copy is usually one
   claim), then add the opening and close. Never pad to a target count or drop a
   supported point to meet one.

## Order

- Give the governing thought before the evidence. Add only the minimum situation and
  complication the reader needs to understand it (SCQA), then the answer.
- An executive summary, when used, maps one row to each body slide or chapter, in the
  same order: one main sentence plus 2–3 details, each from a different angle.
- Each group of peer slides answers one parent question and plays one role (reasons with
  reasons, steps with steps). Choose one order per group: deductive, chronological,
  structural or by importance.
- Give important groups depth and light ones a light touch. Do not make every slide the
  same density or force every list to three items; one point is fine when one suffices.
- End the body with the governing thought at the confidence the evidence supports, and
  the decision or next action you need from the reader. Not "thank you", not a
  restatement. Appendix slides (company profile, hiring, detailed tables) may follow it.
- When facts from before and after a change (a migration, a redesign) sit in one story,
  say on the page which point in time it describes.
- Readers navigate by page: show page numbers, and let the executive summary rows point
  to page ranges (「→ p.15–24」). A source line at the bottom of the slide that uses a
  number or cites another product is part of the slide, not clutter.

## Headlines

Every slide headline is an assertion the reader can agree or disagree with.

- Test 1: can the reader respond "I disagree"? If the only response is "I see what this
  section is about", it is a label. Rewrite it.
- Test 2: hide the body. Does the headline still carry the slide's point? If not, the body
  was carrying the argument.
- Exceptions: cover, section entrances and company introductions use plain labels
  (「会社概要」). A company introduction written as a claim becomes self-praise.
- A chart slide states what the chart means (the "so what"), not what it plots.
- Use specific nouns, real names and numbers, and verbs that say what happened
  (伸びた／止めた／上回った), not 「がある」「を示す」「を提供する」.
- Do not claim more than the evidence supports, especially about the reader's own
  organisation or your own. Narrow the claim to what the body proves.
- Put a count in a headline only when the number is the point (「価格は 8 倍開く」). Element
  counts such as 「3 つの理由」「5 段階」 are not claims; state the content instead. When a
  number does appear, it must match the body.
- No topic label before a colon (「現状と課題：〜」). The topic tag lives in a small
  kicker; the headline starts with the subject.
- A headline stands on its own: no 「この／その／ここまで」 pointing at another slide.
- Peer slides in one group share a grammatical form, even when the group is long; this
  overrides the "no mould three times in a row" rule in writing-ja.md for headlines.
  Across groups, forms naturally vary with role (facts, questions, premises); a whole
  deck of 「〜は、…する」 is a template being copied.
- If no assertion can be written, the slide is either not ready (form the claim first)
  or not needed (merge or cut it). A slide that only says "now the financials" is a
  transition that escaped into its own slide.

For Japanese headlines: plain form, never ending in です／ます (the same holds for the
cover note, footnotes and appendix text you write); two rendered lines at
most (about 40 full-width characters per line at title size, product names in Latin
letters count about half); when it wraps, put an explicit line break at a meaning
boundary and never leave 1–3 characters on the last line; a normal sentence with subject
and object, never compressed into a slogan to fit one line. Do not colour parts of a
headline.

## Turning a talk deck into a document

Keep the facts, numbers and history; rebuild the order and wording answer-first.

- Drop agenda, divider and "what comes next" slides; the summary rows and page numbers do
  their job.
- Move what only the speaker notes say into the body when the reader needs it, including
  limitations the speaker had saved for Q&A: a reader cannot ask.
- Put the speaker and company introductions in an appendix unless the reader needs them
  to trust the argument; keep official copy verbatim.
- Recheck numbers the notes flag as "verify before the talk"; a published document keeps
  them longer.

## Body text

- One slide, one message. Bullets list only what is truly parallel; cause and sequence
  are written as a sentence.
- Within one level, bullets end the same way (all noun endings or all verb endings), never
  です／ます. Different levels may use different endings.
- Keep subject and verb: 「ユーザーが手動で作成する」, not 「手動作成」; 「AI が不備を検出し、
  担当者が確認する」, not 「精度向上」. Do not create passive sentences with abstract nouns as
  subjects.
- Each bullet adds a new angle (cost, time, people, decision, means, volume). Delete any
  bullet whose removal loses no information, and anything the reader already knows.
- One concept, one term, throughout the deck. Expand an abbreviation at first use. Put a
  metaphorical term in 「」 and define it once. Do not bring internal jargon into an
  external deck.
- Say from whose viewpoint relative words hold (社内／外部／先方／現場).
- Number or label items only when a later slide refers back to that number.
- Use parentheses sparingly; do not attach a one-line caption that says what the figure
  already shows.
- For automated behaviour, name the human decision next to it: 「AI が改善案を出し、採用は
  担当者が決める」.

## Tables

- Keep only columns that serve the headline's claim; a column invites the comparison the
  reader will make.
- Rows are members of the set the row-axis defines. Do not slip a conclusion or a
  different category in as a row.
- Column headers are concrete nouns that name what the cells hold, not 「事実」「持っているもの」
  and not conclusions.
- Keep each column at one grain and one viewpoint; write 「—」 for non-applicable cells.
- In a table that shows changes, list only what changed.

## Final check for document decks

1. Read the headlines alone, top to bottom (`deck-text.mjs --mode document`). They must form
   one argument with the answer first, no jumps, repeats or conclusions without premises.
2. Every non-exempt headline passes both assertion tests.
3. Numbers and listed items in headlines match their bodies one to one, and numbers, mechanisms and terms agree across
   pages ([review.md](review.md) step 4). The closing summary claims nothing the body
   contradicts.
4. Evaluative words (限定的, 十分, 問題ない) are backed on the same slide.
5. Later slides add new points instead of restating earlier ones.
6. The last slide asks for a decision or action.
7. A fresh-eye review has been run ([review.md](review.md)).
