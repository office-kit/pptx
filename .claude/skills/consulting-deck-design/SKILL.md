---
name: consulting-deck-design
description: Use when generating a dense, professional consulting/government-style slide deck with pptx-kit (McKinsey/BCG/METI-council-caliber — cover, section dividers, executive summary, exhibits, appendix). Invoked when the user asks for a "McKinsey-style", "consulting-grade", or "board deck" presentation, or says the output looks "too generic" / "too much like a template" / "still too simple / thin compared to the real thing". Covers the visual design system, the chart-authoring pitfalls that make output look amateurish, and the verification workflow — not just "use gradients".
---

# Consulting-grade deck design with pptx-kit

A deck that is schema-valid and opens without repair is a _correct_ deck. It is not
automatically a deck that reads as real consulting/government work. The gap between
"technically fine PowerPoint" and "looks like it came out of a real engagement" is
almost entirely in restraint: fewer decorative flourishes, fewer colors, more
whitespace, and text that states a conclusion instead of a topic — **and** in the
density of substantiation per exhibit (analysis columns, footnotes, superscript
citations) that a first pass tends to skip.

This skill is written from hands-on comparison against two real McKinsey
deliverables:

- **Reference A** — a Japanese METI advisory-council submission (cover, ~10
  exhibits, appendix). Terse, internal-working-deck register: minimal chrome, no
  rule under the headline, action-title-only.
- **Reference B** — a Japanese-language, 68-page public COVID-19 client briefing.
  Richer, external-publication register: rule-under-headline as the default,
  serif headlines, a letter-prefixed "governing thought" convention, four-column
  exhibits, pull-quotes, chevron timelines, and a numbered TOC spine reused per
  section.

Treat every rule below as "here's what we saw break (or land) when compared side by
side with the real thing," not a style preference. Where the two references
disagree (rule-under-headline being the clearest case), the section below says so
explicitly and gives a decision rule instead of picking one silently.

## Before you start: get a reference if you can

If the user has (or can share) an actual example of the target firm's / client's
deck, **read it first** — page through every slide. A generic "corporate blue +
gold" palette invented from scratch reads as generic-corporate, not
McKinsey-specific. Real firm decks have surprisingly little color (near-monochrome
brand-blue family + semantic red/green), and surprisingly little decorative
flourish. If you don't have a reference, default to the palette in this skill
rather than picking your own "premium" colors — gold/orange reads as generic
luxury branding, not strategy consulting.

## Visual language

### Palette

Deep navy → bright blue, plus one teal accent. No gold. Semantic red (risk /
negative) and green (positive) are fine even inside an otherwise blue-monochrome
palette — real decks use them for exactly that purpose, but **sparingly**: across
68 pages of Reference B, red appears on exactly two slides (a legal-disclaimer
line and one "increasing" trend arrow) — it is not a general-purpose "bad news"
data color, it is reserved for the single most alarming signal on the page.

```
NAVY_DARK    #0A1628   — darkest gradient stop, divider/cover/TOC background
NAVY         #132A4D   — solid navy for boxes (org-chart nodes, roadmap headers)
BLUE         #1E4FFF   — primary accent, gradient bright stop, "highlighted" data
TEAL         #00B8D9   — secondary accent
ACCENT_LIGHT #5B9EFF   — tertiary tint for a 4th KPI card / chart series
SLATE        #7A8699   — neutral / "current state" / de-emphasized bars
PANEL_GRAY   #E8E8E8   — light-gray secondary surface (see below)
RED          #D64545   — reserved for risk/negative data points and disclaimers —
                          use on at most one exhibit per section, not as a
                          recurring category color
GREEN        #2E9E6C   — positive, recommended, upside
INK          #111827   — headline / body text (not pure black)
MUTED        #5B6673   — captions, footers, eyebrow labels
```

### Three background modes, used deliberately

Both references converge on exactly three slide-background treatments, each with a
distinct job. Don't invent a fourth (e.g. a mid-blue panel) — it reads as
decoration rather than information hierarchy.

- **White** — the default surface for analytical/exhibit slides (charts, tables).
- **Light gray** (`PANEL_GRAY`, distinctly lighter than white but clearly not
  white) — a secondary "commentary" surface: right-margin "so what" columns,
  contrast callout boxes, a full-panel implications sidebar. Never the primary
  background of a whole deck section — it marks "supporting", not "main".
- **Dark navy** (`NAVY_DARK`/`NAVY`) — cover, section dividers, the TOC/agenda
  spine, "governing thought" framework slides, and (occasionally) a full
  commentary sidebar instead of gray, when the slide wants more visual weight on
  the takeaway than on the exhibit.

A two-panel slide (e.g. white exhibit + gray or navy sidebar) usually needs **no
rule line at the seam at all** — the hard color-block edge already reads as a
divider. Save rule lines for dividing regions that share the same background
(see below).

### Cover and section-divider slides: gradient, not flat fill

Flat navy reads as a corporate template. A real cover uses a diagonal navy→blue
gradient across the full slide. In pptx-kit:

```ts
const bg = addSlideShape(slide, {
  preset: 'rect',
  x: inches(0),
  y: inches(0),
  w: PAGE_W,
  h: PAGE_H,
});
setShapeGradientFill(bg, {
  stops: [
    { offset: 0, color: NAVY_DARK },
    { offset: 1, color: BLUE },
  ],
  angleDeg: 15, // shallow diagonal — the bright stop concentrates toward the right
});
```

Layer white/light-blue title text on top, generous vertical whitespace (the title
block sits roughly at the vertical center, not crammed to the top).

### Headlines: serif, and (usually) a rule underneath

Both references use a **bold serif** typeface for headlines — visibly different
letterforms from the sans-serif body/bullets/table text beneath. This
serif/sans split is one of the highest-leverage, lowest-effort fidelity wins
available, and it was previously unreachable: `setShapeRunFormat`'s `font` field
only sets the _Latin_ typeface (`<a:latin>`); Japanese glyphs render through a
separate East Asian typeface (`<a:ea>`) that has no per-run override in older
pptx-kit versions. Use the current API's `fontEastAsian` field to set both:

```ts
setShapeRunFormat(headlineBox, 0, 0, {
  font: 'Georgia', // Latin glyphs (ASCII digits, English loanwords)
  fontEastAsian: '游明朝', // Yu Mincho — CJK glyphs, serif
  bold: true,
  size: 19,
  color: INK,
});
// Body/bullets/table text stays sans:
setShapeRunFormat(bodyBox, i, 0, { font: 'Arial', fontEastAsian: 'メイリオ', size: 13 });
```

Reserve the serif treatment for: headlines, the "Contents" TOC spine word, named
framework/proper-noun keywords when they appear as section labels (e.g. an
English loanword like "Resolve" used as a chapter name), and nothing else — real
decks never use serif for body copy or data labels.

**Rule-under-headline — decision rule.** Reference A (terse internal deck) never
draws a rule under the title. Reference B (richer external deck) draws a
full-width thin horizontal rule directly beneath the headline on nearly every
content slide — it is the default there, not an exception. For a general-purpose
"McKinsey-caliber" deck, **default to drawing the rule** (`addSlideLine` full
content width, `LINE` color, ~0.75pt) — it's the more consistent convention across
the larger sample and it gives long decks a stronger sense of structure. Drop it
only for the terser register the user explicitly asked for, or on slides where a
hard color-block seam already does the dividing job (two-panel layouts, full-bleed
navy framework slides).

**Letter-prefixed "governing thought" headlines.** For a slide making one load-
bearing claim (the kind of exhibit a reader might photograph out of context),
prefix the headline with a single bold letter + colon in the accent blue,
sequenced A, B, C… across a chapter's exhibits (mirrors the `pointColors` palette
convention, not literal exam grades): `"A: 東南アジア市場は今後5年間、国内の7倍の
ペースで成長する見込み"`. On a white background render the letter in `BLUE`; on a
full-navy background render it the same white/light color as the rest of the
headline (a colored letter loses contrast and reads as an error, not an accent, on
a dark surface — this is a real inconsistency the reference deck itself
resolves that way, not an oversight to "fix"). Reserve the prefix for a chapter's
primary exhibits; supporting/detail slides in between keep a plain headline with
no letter.

If the deck is long enough to need running navigation, keep a small kicker label
above the headline (~9–10pt, gray, regular weight, unbolded) — a quiet wayfinding
aid, not a second headline.

### Numbered TOC / running agenda spine

A distinctive, highly reusable archetype from Reference B: a narrow (~15% width)
white left strip carries a single word (e.g. "Contents") in bold serif, rotated
90° so it reads bottom-to-top like a book spine (`setShapeTextBodyRotationDeg`).
The remaining ~85% is a full-bleed dark-navy field holding one large bold serif
numeral per chapter (`01`, `02`, `03`…), each with a thin rule and a 1–2 line
label beneath it. The **current** chapter's numeral, rule, and label render in
bright accent blue / white; every other chapter renders in muted gray. Reuse this
exact slide once per chapter (only the "active" index changes) — this is the
"running roadmap" convention real long decks use to keep a 60–100-page document
navigable, and it reads as far more deliberate than a single static agenda slide
shown once.

### Left-sidebar "framework" layout

For a slide whose entire point is one short, bold statement plus 3–4 supporting
bullets (e.g. "this strategic pillar," "this technology map"), use a dark
navy-gradient left rail (~28% width) carrying the bold white statement, with the
bullets in the light area to the right. This is a distinct, recognizable layout
from the header+body formula every other slide type uses — real decks reserve it
for framework/thesis slides specifically, not everything.

**Multi-line text formatting gotcha**: if the statement has embedded `\n` line
breaks, each line becomes its own paragraph — `setShapeRunFormat(box, 0, 0, {...})`
only formats paragraph 0. Loop over every paragraph index or the later lines fall
back to the text box's small default style while only the first line looks styled.

### Multi-column exhibits with vertical rule dividers

Reference B's densest, most "real" pages are 3–4 column layouts separated by thin
vertical rule lines (`addSlideLine`, vertical, `LINE` color, ~0.5–0.75pt),
spanning the full content height. Common column contents: a chart flanked by two
text commentary columns; four icon-badge + header + body columns (e.g. a
"5 R's"-style framework); or a bulleted list column beside a pull-quote column.
This is a different device from the single right-margin analysis column already
in this skill (still valid, still the right call for a plain chart-plus-commentary
slide) — reach for the multi-column-with-rules version when a slide has 3+
genuinely parallel blocks to show side by side, not just one exhibit plus one
comment.

### Pull-quote blocks

A quotation callout: a small circular badge (solid `NAVY` fill) containing a
large bold quotation-mark glyph (`"` / `"`, U+201C/U+201D, oversized ~20–24pt,
white or `PANEL_GRAY`), positioned above or beside 1–2 sentences of quoted text,
followed by a bold attribution line introduced with an em-dash
(`— 氏名、肩書き`). Use for a single expert-quote moment per section, not on every
slide.

### Icon strategy (no custom vector art)

Real decks use simple monochrome line-art pictograms (a person, a chart, a
lightbulb, a shield). pptx-kit has no custom-path/freeform-geometry authoring API
and no bundled icon set, so **don't attempt hand-drawn vector icons** — the
practical, reliable substitute is a colored circle or rounded-rect "badge"
(`addSlideShape` `ellipse`/`roundRect`) containing a single safe glyph character,
centered, sized large (~18–24pt). Stick to Latin-1/general-punctuation glyphs —
multi-byte emoji/pictograph ranges are the same CJK-tofu-box risk this skill
already warns about for text, since LibreOffice/PowerPoint font substitution for
emoji is inconsistent across renderers.

```
growth / increase   ▲ or ↑        decrease / decline     ▼ or ↓
success / on-track  ✓             risk / off-track       ! or ✗
insight / idea      ★             direction / next step  →
process step N      a filled circle badge with the bold numeral itself
quote               " / " (U+201C / U+201D), oversized
```

Badge + glyph + header + body is one composite — `groupShapes` it, same as a KPI
card.

### Chart technique: highlight one, gray the rest

The single most-repeated chart device across Reference B: a bar/column series
where **one** category is the narrative subject (bright `BLUE`) and every other
category is a neutral, de-emphasized fill (`SLATE` or a light gray) — not a
categorical rainbow palette. This is authored the same way the pie/doughnut
per-slice override already in this skill works: `pointColors` is a general
series-level field, not restricted to pie/doughnut —

```ts
series: [
  {
    name: '人口100万人当たりの死者数',
    values,
    pointColors: categories.map((c) => (c === 'スウェーデン' ? BLUE : SLATE)),
  },
];
```

Reserve a full multi-hue categorical palette for charts where every category is
equally the subject (e.g. a segment-share pie); use highlight-one-gray-rest when
the slide's headline names a single subject ("A: スウェーデンは…").

### Small provenance tags

Real exhibits mark themselves "非網羅的" ("non-exhaustive") or "例示"
("illustrative") in a small bordered tag in the corner when the chart/table isn't
meant to be a complete list — a plain bordered rect + 8.5pt text in the corner,
not a big banner. Add this to competitive-landscape maps, case-study callouts,
and anything else that is deliberately a sample rather than an exhaustive list.

## Chart authoring — the mistakes that make output look amateurish

These are the actual defects a first pass produces; each one reads as "AI-generated
template" rather than "hand-built exhibit."

### 1. Pie/doughnut charts need explicit per-slice colors

`varyColors: true` (or pptx-kit's default on pie/doughnut, which already sets it)
only _requests_ that the renderer assign distinct colors per slice. It does not
guarantee every viewer actually cycles through a palette — some render every slice
in the same single color, which makes the chart useless. Always pass explicit
`pointColors` cycling through your accent palette:

```ts
series: [{ ...series[0], pointColors: categories.map((_, i) => PALETTE[i % PALETTE.length]) }];
```

### 2. Every multi-series or pie/doughnut chart needs a visible legend

`ChartSpec.legend` is `undefined` by default (no legend authored at all) — pptx-kit
doesn't add one automatically. Without it, a reader can't tell which color means
which series/category. Add `legend: { position: 'r' }` to any chart with more than
one series, or any pie/doughnut (where "categories" function like series). A
single-series bar/line chart can skip the legend — its category-axis labels already
say what each bar/point is, and a legend there is redundant chrome.

### 3. Never plot two series with mismatched scales on one axis

Revenue in the hundreds and a margin percentage in single digits on the same axis
means the percentage line collapses to a flat line near zero — unreadable.
pptx-kit's `ChartSpec` has no secondary axis. Don't try to fake one. Split into two
independent single-metric charts side by side instead — same slide, same story, each
chart properly scaled.

### 4. Don't overlay hand-drawn annotations on top of a native chart

It's tempting to draw dashed reference lines / inline value labels directly on a
chart (mimicking real decks that mark "2030 (48%)" right on the plot). **Don't try
to compute the plot area's pixel position and overlay shapes on it.** A chart's
internal plot rectangle is inset by however much room its own axis labels, legend,
and title need — that inset is decided by the rendering engine (PowerPoint /
LibreOffice / Keynote all differ) and isn't knowable at authoring time. An overlay
aimed at "where category index 3 should be" will drift out of alignment, sometimes
badly (labels landing on top of the y-axis, or off the right edge past the legend).

Instead, put the annotation **beside** the chart, not on it: chart on the left
~68% of the slide width, a right-margin column of 2–4 substantive "so what"
sentences (one per paragraph, a thin colored accent bar to its left) on the
remaining ~28%. This is what both reference decks' own market-growth exhibits
actually do — the "annotations on the chart" look you're copying is itself mostly
a right-margin (or multi-column) text block, not a plot overlay.

## Footnotes and source citations

Real exhibits substantiate every number, and keep two conventions **visually
separate** rather than merging them into one line:

- **Numbered footnotes** (methodology/assumption notes): a small gray numbered
  list ("1.", "2."…) at the bottom of the slide, each tied to a **superscript
  digit** placed inline at the exact claim it qualifies — in a headline, an axis
  label, or mid-bullet (e.g. `"投資回収期間が18ヶ月と最も短く¹"`). Since the
  simple text-authoring API is one run per paragraph (no supported way to mix a
  normal-weight run and a superscript run within a single paragraph), use the
  literal Unicode superscript digit characters (`¹ ² ³` U+00B9/U+00B2/U+00B3,
  `⁴ ⁵ ⁶…` U+2074+) typed directly into the paragraph text — they render as
  normal glyphs in any font, no run-splitting required, and read as true
  superscript footnote markers.
- **Source line**: a separate line beginning `資料:` (or `Source:`), listing
  citations comma-separated. Don't invent per-word hyperlink styling (blue +
  underline on only some words in the line) — that needs multiple differently-
  formatted runs in one paragraph, which the authoring API doesn't support.
  Style the whole line uniformly instead: either plain `MUTED` gray (the more
  common convention — most citation lines in both references are plain, not
  hyperlinked) or, if you want the "clickable source" look, the whole line in a
  muted blue with underline.

Every chart/table exhibit should carry at least a source line; add the numbered-
footnote block whenever a number depends on a stated assumption (survey sample
size, calculation basis, exchange-rate assumption) — this is what separates a
"real" exhibit from a bare chart. See "Composition and structure" below for how
this pairs with a right-margin/multi-column analysis block.

## Composition and structure

- **Group composite elements.** A KPI card (rounded rect + accent tab + big number
  - label), a roadmap node (header bar + numbered badge + title), or an icon
    badge (circle + glyph + header + body) should be one `groupShapes` unit, not
    four independent shapes a future edit could misalign.
- **A real deck's shape**: cover → confidentiality/handling notice → agenda (as a
  reusable numbered TOC spine, not a single static list) → executive summary
  (headline finding, KPI dashboard, 3-pillar recommendation, financial-impact
  chart, roadmap preview) → numbered sections, each opening on a divider slide (or
  the TOC spine with that chapter active) → appendix (detailed tables, glossary,
  sources, disclaimers) → closing. Long decks (60–100+ slides) are normal for a
  real internal working deck; don't compress everything onto 10 slides just
  because a reference example you saw was short — that one may have been a
  client-facing extract, not the full working deck.
- **Every exhibit earns its slide.** A chart or table is never left to speak for
  itself: pair it with a right-margin or multi-column analysis block (2–4
  substantive, numbers-backed sentences — not generic filler) and a footnote/
  source line. A slide with only a bare chart and a one-line caption reads as a
  first draft, not a finished exhibit — information density (not clutter) is
  what makes a page look like it came from a real engagement.
- **Dense but not cluttered**: 3–5 bullets per slide, one action title, one
  exhibit (or one coherent multi-column composition). If you have two unrelated
  exhibits worth showing, that's two slides, not one crowded slide.

## Verification workflow — do not skip this

1. **XSD schema validation** on every generated slide/chart XML (see this repo's
   `test/lib/expect-schema-valid.ts` pattern) — catches structural corruption.
2. **Render with a real engine and look at it.** pptx-kit's own bundled preview
   renderer does not ship CJK fonts — Japanese (or other non-Latin) text renders as
   tofu boxes in it even when the underlying XML is completely correct. That is a
   preview-renderer limitation, not a bug in your output, but it means you **cannot
   trust the bundled preview for a CJK deck.** Convert with a real office suite and
   look at actual pages before calling the deck done:
   ```sh
   soffice --headless --convert-to pdf --outdir /tmp/out deck.pptx
   pdftoppm -png -r 150 -f <page> -l <page> /tmp/out/deck.pdf /tmp/out/page
   ```
   Then view the PNG. Do this for at least one slide of every distinct archetype
   you built (cover, divider, TOC spine, chart, pie chart, table, matrix, sidebar
   layout, multi-column exhibit, pull-quote) — each is a different rendering code
   path and bugs are archetype-specific (e.g. a legend bug affects every pie chart
   but not a single-series bar chart; a footer-position bug affects every
   sidebar-layout slide but not plain content slides).
3. **Read what you rendered critically**, the way the user will: does the pie
   chart's legend actually list distinct colors? Does the multi-line sidebar title
   render at a consistent size on every line? Is the "annotation" text actually
   next to the chart, not drifted onto the axis? Does a two-line footnote block
   fit inside the slide bounds, or does the second line clip off the bottom edge?
   Fix what you find and re-render — don't ship the first render un-inspected.
