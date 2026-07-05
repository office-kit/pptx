---
name: office-kit-pptx
description: Author and edit PowerPoint (.pptx) files from TypeScript/JavaScript with @office-kit/pptx — in Node or the browser. Use when an agent must generate a presentation from scratch, fill a template, or programmatically edit slides, shapes, text, tables, charts, and images, and needs the output to open cleanly in PowerPoint, Keynote, Google Slides, and LibreOffice.
---

# Authoring PPTX with @office-kit/pptx

`@office-kit/pptx` generates **schema-valid** Office Open XML PresentationML. Every
authoring call maps to a specific ECMA-376 element, so a deck you build here
opens and is **fully editable** in PowerPoint — not a flattened image, and not
"valid enough to usually open." That is the bar: _if a file can't be opened and
edited in PowerPoint, it isn't done._

This guide is the fast path for an LLM agent. It covers the canonical call for
each capability, the design rules that separate a real deck from an
AI-template-looking one, the handful of footguns worth memorizing, and a QA
loop you must run before declaring success.

## When to use this

- **Create from scratch** — build a deck programmatically and emit `.pptx`.
- **Fill a template** — load an existing `.pptx`, replace text/images, add
  slides from its layouts, emit the result.
- **Edit** — mutate shapes, text, tables, charts on existing slides.

If you need a _pixel-perfect render_ (print/archival) or format conversion
(Keynote/ODP), this is the wrong tool — use PowerPoint/LibreOffice headless.

## Setup

```sh
npm install @office-kit/pptx
```

```ts
import {
  createPresentation,
  addTitleSlide,
  addContentSlide,
  addBlankSlide,
  addSlideTextBox,
  addSlideShape,
  addSlideTable,
  addSlideChart,
  addSlideImage,
  findSlidePlaceholder,
  setShapeText,
  setShapeRunFormat,
  setShapeFill,
  savePresentation,
} from '@office-kit/pptx';
// Node-only convenience (reads/writes files):
import { loadPresentationFile, savePresentationToFile } from '@office-kit/pptx/node';
```

`savePresentation(pres)` returns a `Uint8Array`. In Node, write it with
`fs.writeFile`; in the browser, wrap it in a `Blob`.

## Mental model (read this once)

1. **One free-function API.** Every capability is a named export that takes the
   thing it operates on (`PresentationData`, `SlideData`, a shape handle) as its
   first argument. There are no classes to instantiate and no fluent chains.
   Import only what you use.
2. **Units are EMU.** Positions/sizes are English Metric Units (914400 per inch).
   Always go through the unit helpers — `inches(1)`, `cm(2)`, `mm(5)`, `pt(18)`,
   `emu(n)` — never raw numbers.
3. **Build, then format.** Add a shape/slide; it returns a handle. Pass that
   handle to formatting functions (`setShape*`). Order of formatting calls does
   not matter — the library inserts each XML child at its schema-mandated slot.
4. **Colors** are `#RRGGBB`, the 3-digit shorthand `#RGB`, bare `RRGGBB`, or a
   theme token (`accent1`…`accent6`, `tx1`, `bg1`, `dk1`, `lt1`, `hlink`). An
   unrecognized color **throws** — it is never silently emitted. One exception:
   chart series colors accept the hex forms but **not** theme tokens (a series
   must resolve to a concrete sRGB value).

## Core workflow — build a deck from scratch

`createPresentation()` returns an immediately-authorable deck (slide master,
Office theme, and `Blank` / `Title Slide` / `Title and Content` layouts) with
no slides. Defaults to 16:9; pass `{ size: '4:3' }` for the classic ratio.

```ts
const pres = createPresentation();

// Title slide — sugar picks the right layout by its locale-stable type token.
const cover = addTitleSlide(pres, 'FY26 Business Review');
const subtitle = findSlidePlaceholder(cover, 'subTitle');
if (subtitle) setShapeText(subtitle, 'Strategy, results, and the road ahead');

// Title + body content slide.
const agenda = addContentSlide(pres, { title: 'Agenda' });
const body = findSlidePlaceholder(agenda, 'body');
// IMPORTANT: bullet *content* is multi-line text + a bullet style, NOT a list arg.
if (body) setShapeText(body, 'Highlights\nFinancials\nRoadmap\nRisks', { bullets: 'bullet' });

const out = await savePresentation(pres);
```

Slide constructors:

| Goal                              | Call                                                              |
| --------------------------------- | ----------------------------------------------------------------- |
| Title slide                       | `addTitleSlide(pres, title)`                                      |
| Title + body (bulleted)           | `addContentSlide(pres, { title, body })`                          |
| Section divider                   | `addSectionHeaderSlide(pres, title)`                              |
| Empty canvas for free-form layout | `addBlankSlide(pres)`                                             |
| Bind a layout explicitly          | `addSlide(pres, { layout: findSlideLayoutByType(pres, 'obj')! })` |

Prefer `findSlideLayoutByType(pres, 'title' | 'obj' | 'secHead' | 'blank')` —
the `type` token is stable across PowerPoint UI languages. `findSlideLayout`
matches the localized, case-sensitive display name.

## Capability cheat-sheet (the canonical call)

Content on a blank slide (all positions via `inches`/`cm`/`pt`):

```ts
const slide = addBlankSlide(pres);

// Text box.
const tb = addSlideTextBox(slide, {
  x: inches(0.6),
  y: inches(0.4),
  w: inches(8),
  h: inches(1),
  text: 'A strong year',
});
setShapeRunFormat(tb, 0, 0, { bold: true, size: 32, color: '#1F2937', font: 'Arial' });

// Preset shape with centered text + gradient + shadow.
const card = addSlideShape(slide, {
  preset: 'roundRect',
  x: inches(0.6),
  y: inches(1.6),
  w: inches(3.6),
  h: inches(2.2),
  text: 'Revenue +38%',
  textAnchor: 'ctr',
});
setShapeGradientFill(card, {
  stops: [
    { offset: 0, color: '#2563EB' },
    { offset: 1, color: '#1E3A8A' },
  ],
  angleDeg: 90,
});
setShapeShadow(card, {
  color: '#000000',
  blurEmu: pt(8),
  offsetEmu: pt(3),
  angleDeg: 90,
  opacity: 0.35,
});

// Table (firstRow + bandRow give the banded header look).
addSlideTable(slide, {
  x: inches(0.6),
  y: inches(1.6),
  w: inches(8.2),
  h: inches(2.5),
  rows: [
    ['Metric', 'FY25', 'FY26'],
    ['Revenue', '$120M', '$166M'],
    ['Margin', '64%', '67%'],
  ],
  firstRow: true,
  bandRow: true,
});

// Chart (embedded workbook + caches are generated automatically).
addSlideChart(slide, {
  x: inches(0.6),
  y: inches(1.1),
  w: inches(8.2),
  h: inches(4.2),
  spec: {
    kind: 'column', // bar | column | line | pie | doughnut | area
    categories: ['Q1', 'Q2', 'Q3', 'Q4'],
    series: [
      { name: 'Revenue', values: [120, 138, 152, 166] },
      { name: 'Plan', values: [115, 130, 148, 160] },
    ],
    title: 'Revenue vs plan ($M)',
    dataLabels: { showValue: true, showCategory: false, showSeriesName: false, showPercent: false },
  },
});

// Image (format auto-detected from the bytes; `fit` controls letterbox/crop).
addSlideImage(slide, pngBytes, {
  x: inches(0.6),
  y: inches(1.1),
  w: inches(8.2),
  h: inches(4.2),
  fit: 'contain',
});
```

Formatting and slide features (one canonical call each):

| Capability                      | Call                                                                                                                       |
| ------------------------------- | -------------------------------------------------------------------------------------------------------------------------- |
| Whole-shape text                | `setShapeText(shape, text, { bullets? })` (split lines with `\n`)                                                          |
| One run's format                | `setShapeRunFormat(shape, p, r, { bold, italic, underline, size, color, font, highlight, ... })`                           |
| Paragraph align / level         | `setParagraphAlignment(shape, p, 'ctr')`, `setParagraphLevel(shape, p, 1)`                                                 |
| Paragraph spacing / leading     | `setParagraphSpacing(shape, p, { beforePts, afterPts })`, `setParagraphLineSpacing(shape, p, { kind: 'pct', value: 1.5 })` |
| Solid / gradient / pattern fill | `setShapeFill(shape, '#2563EB')`, `setShapeGradientFill(...)`, `setShapePatternFill(...)`                                  |
| Outline                         | `setShapeStroke(shape, { color, widthEmu })` + `setShapeStrokeDash/Arrow/Cap/Join`                                         |
| Effects                         | `setShapeShadow(shape, {...})`, `setShapeGlow(shape, {...})`                                                               |
| Geometry                        | `setShapePosition/Size/Rotation/Flip/Bounds`, `bringShapeToFront`, `sendShapeToBack`                                       |
| Picture corrections             | `setShapeImageCrop/Opacity/Brightness/Contrast` (brightness/contrast in `[-1, 1]`)                                         |
| Hyperlink / click action        | `setShapeHyperlink(shape, url)`, `setShapeClickAction(shape, { kind: 'nextSlide' })`                                       |
| Slide background                | `setSlideBackground(slide, '#102030')`, `setSlideBackgroundImage(slide, bytes)`                                            |
| Transition                      | `setSlideTransition(slide, { effect: 'fade' })` — key is **`effect`**, not `type`                                          |
| Animation                       | `setShapeAnimation(shape, { effect: 'fadeIn' })` (`fadeIn`/`fadeOut`/`appear`/`disappear`)                                 |
| Speaker notes                   | `setSlideNotes(slide, '...')`                                                                                              |
| Comments                        | `addSlideComment(slide, { author: { name }, text })`                                                                       |
| Sections                        | `setSlideSections(pres, [{ name, slides: [...] }])`                                                                        |

## Fill a template instead

```ts
import {
  loadPresentation,
  getSlides,
  findSlidePlaceholder,
  setShapeText,
  replaceTokensInPresentation,
  setShapeImage,
  getSlideShapes,
  getShapeKind,
  getShapeName,
  savePresentation,
} from '@office-kit/pptx';

const pres = await loadPresentation(templateBytes);

// Placeholder text:
const title = findSlidePlaceholder(getSlides(pres)[0]!, 'title');
if (title) setShapeText(title, 'Q3 Review');

// Token fill across every slide ({{name}}, {{date}}, ...):
replaceTokensInPresentation(pres, { name: 'Alice', date: '2026-12-01' });

// Swap an image in place (geometry preserved):
for (const s of getSlides(pres))
  for (const sh of getSlideShapes(s))
    if (getShapeKind(sh) === 'picture' && getShapeName(sh) === 'Logo')
      setShapeImage(sh, newLogoBytes);

const out = await savePresentation(pres);
```

Out-of-scope parts are preserved on round-trip — the library never silently
strips content it doesn't model.

## Design rules (so it doesn't look AI-generated)

A schema-valid deck can still look like a template. Before building, commit to a
look and apply it consistently:

- **Pick a bold, content-informed palette** and let _one_ color dominate
  (~60–70% of visual weight), with one accent. Don't scatter six accent colors
  per slide. Define the palette once as hex constants and reuse it.
- **Every slide needs a visual element** — a chart, table, image, colored
  shape, or a strong type hierarchy. Avoid bullet-only text slides; convert
  lists into two-column layouts, icon/number rows, or small-multiple cards.
- **Don't repeat the same layout** slide after slide. Alternate: full-bleed
  title, two-column, chart-led, table-led, quote/stat callout.
- **Typography**: titles 32–44pt bold, body 14–18pt, one or two font families.
  Keep a ≥0.5in (`inches(0.5)`) margin from the slide edge; don't crowd edges.
- **Numbers deserve charts**, not sentences. A trend → line; parts of a whole →
  pie/doughnut; comparison across categories → column/bar.
- **Avoid the tells**: a thin accent line under every title, dense walls of
  bullets, clip-art, and four different accent colors all read as "generated."

Sizing reference (16:9 deck): the canvas is `inches(13.333) × inches(7.5)`.
Keep content within `x ∈ [0.5, 12.83]`, `y ∈ [0.5, 7.0]` inches.

## Footguns (memorize these — each is a real, easy mistake)

- **Bullets are content + style, not a list argument.** To make a bulleted
  list: `setShapeText(shape, 'A\nB\nC', { bullets: 'bullet' })`.
  `setShapeBullets(shape, style)` sets the _bullet glyph style_ (`'bullet'` |
  `'number'` | `'none'` | `{ char }` | `{ autoNum }`) on existing paragraphs —
  it is NOT how you set the text.
- **`setShapeFill(shape, color)` takes a color string**, e.g.
  `setShapeFill(card, '#059669')` — not an object.
- **Transitions key on `effect`**: `setSlideTransition(slide, { effect: 'fade' })`.
  `{ type: 'fade' }` is wrong. `effect: 'none'` emits no transition (use
  `clearSlideTransition` to remove one).
- **Multi-line text** in a text box, shape, or table cell uses `\n` between
  lines — each becomes its own paragraph. A literal newline inside one run is
  not a line break.
- **Authorable chart kinds** are `bar`, `column`, `line`, `pie`, `doughnut`,
  `area`. `scatter`/`radar`/`bubble` are read-only today (authoring them
  throws). `pie`/`doughnut` take exactly one series.
- **Find placeholders by type token**, not display name:
  `findSlidePlaceholder(slide, 'title' | 'body' | 'ctrTitle' | 'subTitle')`.

## QA protocol — run this before saying "done"

1. **Structural validation.** `const issues = validatePresentation(pres);` —
   treat any `severity === 'error'` as a blocker (missing rels, dangling slide
   ids, layouts without masters, etc.). Fix and re-run until clean.
2. **Round-trip.** `loadPresentation(await savePresentation(pres))` must not
   throw and must report the same slide count.
3. **Schema validity.** If you have `xmllint` + the ECMA-376 XSDs, validate each
   emitted part. The library's own test suite gates this; if you author novel
   combinations, validate them too.
4. **Visual check.** Render to an image and _look at it_ — overflowing text,
   collisions, off-canvas shapes, and unreadable color contrast do not show up
   in schema validation. Use `@office-kit/pptx-preview` (SVG in the browser, PNG on the
   server) or open the file in PowerPoint/LibreOffice. `findShapesOutsideCanvas`
   and `findOverlappingShapePairs` catch layout problems programmatically.
5. **Content check.** Grep the saved deck's text (`getPresentationText(pres)`)
   for leftover placeholder tokens (`{{`, "Lorem", "TODO") before shipping.

## Known limitations (don't fight these — design around them)

Read-pass-through or post-1.0, _not_ authorable today:

- `scatter` / `radar` / `bubble` charts, combo charts, and secondary value axes.
- Multiple independently-formatted runs **within one paragraph** (inline rich
  text like "make _this word_ bold"); each paragraph is authored as one run.
  Use per-shape or per-paragraph formatting, or split across shapes.
- Constructing new themes / masters / layouts from scratch (author _on top of_
  the ones `createPresentation` or a template provides).
- SmartArt authoring, complex multi-step animation timing trees, OLE/ActiveX,
  and document encryption.
- Modern threaded comments (legacy `<p:cm>` comments are read + written).

## Worked example

A complete, validated multi-slide business deck is in
[`examples/business-deck.md`](examples/business-deck.md). It is exercised by the
library's test suite (`test/skill-example.test.ts`), so the code there is known
to produce a schema-valid deck.
