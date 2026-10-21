# Visual design

Read this before styling a new deck or restyling an existing one. A deck that follows
the story rules but sets body text at 10 pt on white with one blue reads as a draft:
the reader judges the look before the argument. When revising a deck that already has a
design, keep its palette and type scale and apply only the floors below.

Values assume the default 16:9 slide (13.333 × 7.5 in). Put the palette and type scale
in `theme.ts` and import them; do not pick sizes or colors per slide.

## Type sizes

| Text                                  | Size                            |
| ------------------------------------- | ------------------------------- |
| Slide headline                        | 26–32 pt (talk divider: larger) |
| Body text, card text                  | 14 pt minimum                   |
| Dense table body, chart axis labels   | 12 pt minimum                   |
| Footnotes, source lines, page numbers | 10 pt minimum                   |

These are floors, not targets: a talk deck seen across a room needs 18 pt or more for
body text. When content does not fit at the floor, split the slide, cut words or move
detail to the notes or an appendix. Never shrink below the floor to make it fit.

Use two or three sizes per slide. The headline, the body and the small print must be
distinguishable at a glance; 14 pt body under a 16 pt headline is no hierarchy.

## Cover and closing slides

The cover and the last slide are the ones people remember and forward. Do not leave them
as black text on white.

- Fill the slide, or a large block of it, with the primary color and set the text in
  white (or the other way round for a dark theme).
- Put the title, one line of context (who, when, for what meeting) and, when the deck
  has one, the key figure the deck is about: 「解約率 4.2% → 2.8%」, 「承認額 1,200 万円」.
  Use only figures from the source material.
- The closing slide repeats that figure next to the ask or the next action, in the same
  color field, so the deck visibly ends.

## Color

Use one primary color and one accent color, plus neutrals (text, greys, white).

- The **primary** color carries structure: header rows, the cover field, section marks,
  the main chart series.
- The **accent** color means "this is the problem" or "this is the point": the bar that
  missed the target, the gap, the one number to look at. Use it only for that meaning,
  at most once or twice per slide. If everything is accented, nothing is.
- Context and comparison data are grey. A chart where every series has its own bright
  color asks the reader to compare all of them.
- Check contrast with [the TSX reference](tsx.md) ("Readable colors"); do not rely on
  color alone to carry meaning.

## Layout and cards

- Keep one margin (about 0.5–0.6 in) and one left edge for headlines and body on every
  slide. Align objects to each other; do not place them by eye.
- Cards in one row share one top edge and one height. Size the height to the fullest
  card plus padding, then trim the content of the others; a row with one tall card and
  three half-empty ones looks broken. Keep at least 0.2 in between text and the card
  edge.

## Charts

- The headline states what the chart shows (the "so what"); the chart proves it. Draw
  only the series that serve that claim.
- Put the unit where it is read: a value-axis title (`valueAxisTitle: '件'`), the
  number format (`valueAxis: { numberFormat: '0%' }`) or the category labels. A number
  without a unit is a question.
- Start a bar or column value axis at zero. When a line chart's axis must start
  elsewhere to show the movement, say so on the slide (「縦軸は 60 から」).
- Draw a target or plan as a constant series (`chartKind: 'line'`, the same value in
  every category), not as a `Line` shape laid over the chart: a shape does not move
  when the data or axis changes. Name it in the legend or label its end.
- Label the values the headline talks about with data labels, and leave the rest to the
  axis. Labels on every point compete with the message.
- Color the series in the message with the primary or accent color and the rest grey.

## Tables

- Right-align numbers and keep their decimal places and units consistent within a
  column; left-align text. With `Table`, use `styleCell` to set `align: 'right'` on the
  number columns.
- One header row. Put the unit in the header (「売上（百万円）」) rather than in every
  cell. If you need a two-level header, the table is usually two tables.
- Body text 12 pt minimum and no cell longer than two lines; see
  [the document rules](story-document.md) ("Tables") for long-text columns.
- Use `stripeFill` or thin horizontal rules for long tables; avoid a full grid of
  vertical lines.

## Final check for the look

1. No text below the floors above; nothing overflows, clips or collides.
2. Cover and closing use a color field and carry the key figure when the deck has one.
3. One primary and one accent color; every accent marks a problem or the point.
4. Cards in a row share height and edges.
5. Every chart has its unit, a zero baseline or a note saying it has none, and targets
   drawn as series.
6. Table numbers are right-aligned under a single header row.
