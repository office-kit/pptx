---
name: consulting-deck-design
description: Use when generating a dense, professional consulting/government-style slide deck with pptx-kit (McKinsey/BCG/METI-council-caliber — cover, section dividers, executive summary, exhibits, appendix). Invoked when the user asks for a "McKinsey-style", "consulting-grade", or "board deck" presentation, or says the output looks "too generic" / "too much like a template". Covers the visual design system, the chart-authoring pitfalls that make output look amateurish, and the verification workflow — not just "use gradients".
---

# Consulting-grade deck design with pptx-kit

A deck that is schema-valid and opens without repair is a *correct* deck. It is not
automatically a deck that reads as real consulting/government work. The gap between
"technically fine PowerPoint" and "looks like it came out of a real engagement" is
almost entirely in restraint: fewer decorative flourishes, fewer colors, more
whitespace, and text that states a conclusion instead of a topic.

This skill is written from hands-on comparison against a real McKinsey deliverable
(a Japanese METI advisory-council submission — cover, ~10 exhibits, appendix). Treat
every rule below as "here's what broke when we didn't do this," not a style
preference.

## Before you start: get a reference if you can

If the user has (or can share) an actual example of the target firm's / client's
deck, **read it first** — page through every slide. A generic "corporate blue +
gold" palette invented from scratch reads as generic-corporate, not
McKinsey-specific. Real firm decks have surprisingly little color (near-monochrome
brand-blue family + semantic red/green), and surprisingly little chrome (no boxed
headers, no rules under every title). If you don't have a reference, default to the
palette in this skill rather than picking your own "premium" colors — gold/orange
reads as generic luxury branding, not strategy consulting.

## Visual language

### Palette

Deep navy → bright blue, plus one teal accent. No gold. Semantic red (risk /
negative) and green (positive) are fine even inside an otherwise blue-monochrome
palette — real decks use them for exactly that purpose.

```
NAVY_DARK    #0A1628   — darkest gradient stop, divider/cover background
NAVY         #132A4D   — solid navy for boxes (org-chart nodes, roadmap headers)
BLUE         #1E4FFF   — primary accent, gradient bright stop
TEAL         #00B8D9   — secondary accent
ACCENT_LIGHT #5B9EFF   — tertiary tint for a 4th KPI card / chart series
SLATE        #7A8699   — neutral / "current state" bars
RED          #D64545   — risk, negative, decline
GREEN        #2E9E6C   — positive, recommended, upside
INK          #111827   — headline / body text (not pure black)
MUTED        #5B6673   — captions, footers, eyebrow labels
```

### Cover and section-divider slides: gradient, not flat fill

Flat navy reads as a corporate template. A real cover uses a diagonal navy→blue
gradient across the full slide. In pptx-kit:

```ts
const bg = addSlideShape(slide, { preset: 'rect', x: inches(0), y: inches(0), w: PAGE_W, h: PAGE_H });
setShapeGradientFill(bg, {
  stops: [{ offset: 0, color: NAVY_DARK }, { offset: 1, color: BLUE }],
  angleDeg: 15, // shallow diagonal — the bright stop concentrates toward the right
});
```

Layer white/light-blue title text on top, generous vertical whitespace (the title
block sits roughly at the vertical center, not crammed to the top).

### Content-slide headlines: bold statement, no rule, no boxed kicker

Don't do the "gold all-caps eyebrow + big headline + colored rule underneath"
formula — it's a template signature, not a consulting one. Real slides put a bold,
dark, ~19–20pt **action title** directly near the top with generous whitespace below
and nothing separating it from the body. An "action title" states the *conclusion*
("東南アジア市場は今後5年間、国内の7倍のペースで成長する見込み"), not the topic
("市場成長トレンド") — this matters more than any color choice.

If the deck is long enough to need running navigation, keep a section label, but
make it small (~9–10pt), gray, regular weight, unbolded, uncolored — a quiet
wayfinding aid, not a second headline.

### Left-sidebar "framework" layout

For a slide whose entire point is one short, bold statement plus 3–4 supporting
bullets (e.g. "this strategic pillar," "this technology map"), use a dark
navy-gradient left rail (~28% width) carrying the bold white statement, with the
bullets in the light area to the right. This is a distinct, recognizable layout from
the header+body formula every other slide type uses — real decks reserve it for
framework/thesis slides specifically, not everything.

**Multi-line text formatting gotcha**: if the statement has embedded `\n` line
breaks, each line becomes its own paragraph — `setShapeRunFormat(box, 0, 0, {...})`
only formats paragraph 0. Loop over every paragraph index or the later lines fall
back to the text box's small default style while only the first line looks styled.

### Small provenance tags

Real exhibits mark themselves "非網羅的" ("non-exhaustive") or "例示" ("illustrative")
in a small bordered tag in the corner when the chart/table isn't meant to be a
complete list — a plain bordered rect + 8.5pt text in the corner, not a big banner.
Add this to competitive-landscape maps, case-study callouts, and anything else that
is deliberately a sample rather than an exhaustive list.

## Chart authoring — the mistakes that make output look amateurish

These are the actual defects a first pass produces; each one reads as "AI-generated
template" rather than "hand-built exhibit."

### 1. Pie/doughnut charts need explicit per-slice colors

`varyColors: true` (or pptx-kit's default on pie/doughnut, which already sets it)
only *requests* that the renderer assign distinct colors per slice. It does not
guarantee every viewer actually cycles through a palette — some render every slice
in the same single color, which makes the chart useless. Always pass explicit
`pointColors` cycling through your accent palette:

```ts
series: [{ ...series[0], pointColors: categories.map((_, i) => PALETTE[i % PALETTE.length]) }]
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
~68% of the slide width, a right-margin column of 2–3 short "so what" sentences
(one per paragraph, a thin colored accent bar to its left) on the remaining ~28%.
This is what the real reference deck's own market-growth exhibit actually does —
the "annotations on the chart" look you're copying is itself mostly a
right-margin column, not a plot overlay.

## Composition and structure

- **Group composite elements.** A KPI card (rounded rect + accent tab + big number
  + label) or a roadmap node (header bar + numbered badge + title) should be one
  `groupShapes` unit, not four independent shapes a future edit could misalign.
- **A real deck's shape**: cover → confidentiality/handling notice → agenda →
  executive summary (headline finding, KPI dashboard, 3-pillar recommendation,
  financial-impact chart, roadmap preview) → numbered sections, each opening on a
  divider slide → appendix (detailed tables, glossary, sources, disclaimers) →
  closing. Long decks (60–100+ slides) are normal for a real internal working deck;
  don't compress everything onto 10 slides just because the reference example you
  saw was short — that one was a client-facing extract, not the full working deck.
- **Dense but not cluttered**: 3–5 bullets per slide, one action title, one
  exhibit. If you have two exhibits worth showing, that's two slides (or a
  chart+table split slide), not one crowded slide.

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
   you built (cover, divider, chart, pie chart, table, matrix, sidebar layout) —
   each is a different rendering code path and bugs are archetype-specific (e.g. a
   legend bug affects every pie chart but not a single-series bar chart).
3. **Read what you rendered critically**, the way the user will: does the pie
   chart's legend actually list distinct colors? Does the multi-line sidebar title
   render at a consistent size on every line? Is the "annotation" text actually
   next to the chart, not drifted onto the axis? Fix what you find and re-render —
   don't ship the first render un-inspected.
