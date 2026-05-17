# pptx-kit

Generate and edit `.pptx` (PowerPoint / Office Open XML Presentation) files
from TypeScript — in **Node.js or the browser**, from a single ESM bundle.

> **Status: 1.0 — public API stabilized.** Every capability in the table
> below works end-to-end against real PPTX fixtures, with every emitted XML
> part validated against the ECMA-376 schemas via `xmllint` in CI. Future
> 1.x releases are SemVer-compatible.

## Why

The JavaScript ecosystem has several PPTX libraries, but they typically pick
one trade-off:

- **Node-only** with a Buffer-shaped API → does not work in the browser.
- **Browser-only** wrapping a fixed template → cannot author from scratch.
- **Loose XML strings** that "usually open" → break in Keynote / Google Slides
  / the Open XML SDK validator.

`pptx-kit` is built around a different stance:

- One ESM bundle that runs in **Node and the browser**.
- A typed object model that mirrors the **OOXML PresentationML** spec
  (ECMA-376 Part 1, §19). When the spec says something is a choice, our types
  say it is a discriminated union.
- Output that passes Microsoft's
  [Open XML SDK Productivity Tool](https://github.com/dotnet/Open-XML-SDK)
  validator, not just PowerPoint's "open and pray."
- Two complementary paths: **author from scratch** _or_ **edit a template**.

## Scope

The work is split into four levels of completeness. The v1.0 release targets
levels 1-3 in full and level 4 in part:

| Level | Capability                                                          | v1.0                            |
| ----- | ------------------------------------------------------------------- | ------------------------------- |
| L1    | Read an existing PPTX, save it back without corruption              | ✅                              |
| L2    | Template edit — text replacement, image swap, add slide from layout | ✅                              |
| L3    | Authoring — shapes, text, tables, fills, effects, transforms        | ✅                              |
| L3    | Authoring on top of existing themes / masters / layouts             | ✅                              |
| L3    | Constructing new themes / masters / layouts from scratch            | ❌ post-1.0                     |
| L3    | Charts (all common types) with embedded data                        | ✅                              |
| L4    | Notes, comments, transitions                                        | ✅                              |
| L4    | Simple animations (entrance / exit / emphasis presets)              | ✅                              |
| L4    | SmartArt authoring                                                  | ❌ post-1.0 (read pass-through) |
| L4    | Complex animation timing trees                                      | ❌ post-1.0                     |
| L4    | OLE / ActiveX authoring                                             | ❌ post-1.0 (read pass-through) |
| L4    | Document encryption (read + write)                                  | ❌ post-1.0                     |

Out-of-scope content is still **preserved on round-trip** — `pptx-kit` will
never silently strip parts it doesn't model. That's the L1 contract.

When NOT to use this:

- You want to **render** PPTX to pixels / PDF in the browser. Use a renderer
  (e.g. `pptx2html`, server-side LibreOffice headless). `pptx-kit` writes
  PPTX, it does not paint it.
- You need a thin DSL for one-off "report" slides and do not care about
  schema validity. A simpler library will be lighter.
- You want to convert PPTX to another format (Keynote, ODP). Out of scope
  forever — that's a renderer's job.

## Install

```sh
npm install pptx-kit
# or
pnpm add pptx-kit
# or
yarn add pptx-kit
```

## One API

pptx-kit exposes a single tree-shakeable free-function API. Every
capability is a named export — `loadPresentation`, `savePresentation`,
`addSlideTextBox`, `setShapeFill`, etc. Bundlers drop every entry you
don't import, so the minimal `load → save` bundle is **~60 KB**.

```ts
import {
  findSlidePlaceholder,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeText,
} from 'pptx-kit';

const pres = await loadPresentation(bytes);
const title = findSlidePlaceholder(getSlides(pres)[0]!, 'title');
if (title) setShapeText(title, 'Hello');
const out = await savePresentation(pres);
```

CI enforces the tree-shake bound in `test/tree-shake.test.ts`.

## Usage

### Edit a template

```ts
import {
  findSlidePlaceholder,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeText,
} from 'pptx-kit';

const pres = await loadPresentation(existingPptxBytes);
const cover = getSlides(pres)[0]!;
const title = findSlidePlaceholder(cover, 'title');
if (title) setShapeText(title, 'Q3 Review');
const body = findSlidePlaceholder(cover, 'body');
if (body) setShapeText(body, 'Numbers up and to the right.');

const out: Uint8Array = await savePresentation(pres);
// Node:    fs.writeFile('out.pptx', out)
// Browser: new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' })
```

### Token-based template fill

```ts
import { loadPresentation, replaceTokensInPresentation, savePresentation } from 'pptx-kit';

const pres = await loadPresentation(templateBytes);
// Replaces `{{name}}`, `{{event}}`, `{{date}}` across every slide.
replaceTokensInPresentation(pres, { name: 'Alice', event: 'Re:Invent', date: '2026-12-01' });
const out = await savePresentation(pres);
```

### Build a deck from a blank template

```ts
import {
  addSlide,
  addSlideImage,
  addSlideTextBox,
  duplicateSlide,
  findSlideLayout,
  findSlidePlaceholder,
  inches,
  loadPresentation,
  moveSlide,
  savePresentation,
  setShapeText,
} from 'pptx-kit';

const pres = await loadPresentation(await fetch('/blank.pptx').then((r) => r.arrayBuffer()));

const titleLayout = findSlideLayout(pres, 'Title Slide')!;
const slide1 = addSlide(pres, { layout: titleLayout });
setShapeText(findSlidePlaceholder(slide1, 'ctrTitle')!, 'pptx-kit demo');
setShapeText(findSlidePlaceholder(slide1, 'subTitle')!, 'an OOXML library for TypeScript');

const blank = findSlideLayout(pres, 'Blank')!;
const slide2 = addSlide(pres, { layout: blank });
addSlideTextBox(slide2, {
  x: inches(1),
  y: inches(1),
  w: inches(8),
  h: inches(1),
  text: 'Free-form text box',
});
addSlideImage(slide2, imageBytes, { x: inches(1), y: inches(3), w: inches(3), h: inches(3) });

const dup = duplicateSlide(pres, slide2);
moveSlide(pres, dup, 0);

const out: Uint8Array = await savePresentation(pres);
```

### Replace an image in place

```ts
import {
  getShapeKind,
  getShapeName,
  getSlideShapes,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeImage,
} from 'pptx-kit';

const pres = await loadPresentation(templateBytes);
for (const slide of getSlides(pres)) {
  for (const shape of getSlideShapes(slide)) {
    if (getShapeKind(shape) === 'picture' && getShapeName(shape) === 'Logo') {
      setShapeImage(shape, newLogoBytes); // format auto-detected; geometry preserved
    }
  }
}
const out = await savePresentation(pres);
```

### Node convenience entry

```ts
import { loadPresentationFile, savePresentationToFile } from 'pptx-kit/node';

const pres = await loadPresentationFile('./template.pptx');
await savePresentationToFile(pres, './out.pptx');
```

### Charts

```ts
import { addSlideChart, getSlides, loadPresentation, savePresentation, inches } from 'pptx-kit';

const pres = await loadPresentation(templateBytes);
const slide = getSlides(pres)[0];
addSlideChart(slide!, {
  x: inches(0.5),
  y: inches(0.5),
  w: inches(8),
  h: inches(4.5),
  spec: {
    kind: 'column', // bar | column | line | pie | doughnut | area
    categories: ['Q1', 'Q2', 'Q3', 'Q4'],
    series: [
      { name: 'Revenue', values: [120, 180, 240, 300] },
      { name: 'Cost', values: [80, 90, 130, 160] },
    ],
    title: 'FY26 plan',
  },
});

await savePresentation(pres);
```

The embedded xlsx that PowerPoint requires for "Edit data" is generated
automatically. Inline `<c:strCache>` / `<c:numCache>` caches mean the
chart renders without opening the workbook.

### Animations

```ts
import { setShapeAnimation, getSlideShapes, getSlides } from 'pptx-kit';

const slide = getSlides(pres)[0]!;
const shape = getSlideShapes(slide)[0]!;
setShapeAnimation(shape, { effect: 'fadeIn', durationMs: 800 });
// effects: 'fadeIn' | 'fadeOut' | 'appear' | 'disappear'
```

### Comments

```ts
import { addSlideComment, getSlides } from 'pptx-kit';

const slide = getSlides(pres)[0]!;
addSlideComment(slide, {
  author: { name: 'Reviewer A' },
  text: 'Punch up the numbers here.',
  position: { x: 1_000_000, y: 1_000_000 }, // optional EMU coords
});
```

### Gradient fills

```ts
import { setShapeGradientFill } from 'pptx-kit';

setShapeGradientFill(shape, {
  stops: [
    { offset: 0, color: '#FF0000' },
    { offset: 1, color: '#0000FF' },
  ],
  angleDeg: 90, // top → bottom
});
```

### Validation

```ts
import { validatePresentation } from 'pptx-kit';

const issues = validatePresentation(pres);
for (const i of issues) console.error(i.severity, i.message);
// Catches missing rels, dangling slide ids, layouts without masters, etc.
```

### API surface (current state)

Each row lists the free-function entry points. Read/write pairs are
shown together.

| Capability           | API                                                                                                                                                                                                                                                     |
| -------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Load / save          | `loadPresentation(input)`, `savePresentation(pres)`, `loadPresentationFile(path)` (node), `savePresentationToFile(pres, path)` (node)                                                                                                                   |
| Create               | `createPresentation()`                                                                                                                                                                                                                                  |
| Slide CRUD           | `getSlides`, `getSlideAt`, `getSlideIndex`, `addSlide`, `removeSlide`, `moveSlide`, `duplicateSlide`, `clearSlideShapes`                                                                                                                                |
| Slide layout         | `getSlideLayouts`, `findSlideLayout`, `getSlideLayout(slide)`, `setSlideLayout(slide, layout)`, `getSlideLayoutName`, `getSlideLayoutType`                                                                                                              |
| Slide metadata       | `getSlideTitle` / `setSlideTitle`, `getSlideSize` / `setSlideSize`, `isSlideHidden` / `setSlideHidden`, `getSlideText`                                                                                                                                  |
| Slide sections       | `getSlideSections`, `setSlideSections` (p14 sectionLst)                                                                                                                                                                                                 |
| Placeholders         | `findSlidePlaceholder(slide, 'title' \| 'body' \| ...)`                                                                                                                                                                                                 |
| Token / text replace | `replaceTokensInPresentation`, `replaceTokensInSlide`, `replaceTextInPresentation`, `replaceTextInSlide`                                                                                                                                                |
| Background           | `getSlideBackground` / `setSlideBackground` / `clearSlideBackground`                                                                                                                                                                                    |
| Notes                | `getSlideNotes` / `setSlideNotes`                                                                                                                                                                                                                       |
| Transitions          | `getSlideTransition` / `setSlideTransition` / `clearSlideTransition`                                                                                                                                                                                    |
| Animations           | `getShapeAnimation` / `setShapeAnimation` (`fadeIn` / `fadeOut` / `appear` / `disappear`), `clearSlideAnimations`                                                                                                                                       |
| Comments             | `addSlideComment`, `getSlideComments`, `removeSlideComment`, `getCommentAuthors`, `getCommentText` / `getCommentAuthor` / `getCommentPosition`                                                                                                          |
| Shape authoring      | `addSlideTextBox`, `addSlideShape`, `addSlideLine`, `addSlideTable`, `addSlideImage`, `addSlideChart`                                                                                                                                                   |
| Shape lookup         | `findShapeByName`, `findShapesByName`, `findShapesByKind`, `findShapeInPresentation`, `getAllShapes`, `getSlideShapes`                                                                                                                                  |
| Shape text           | `setShapeText`, `setShapeBullets`, `setShapeAlignment`, `setShapeTextFormat`, `setShapeHyperlink` / `getShapeHyperlink`                                                                                                                                 |
| Per-paragraph        | `setParagraphAlignment` / `getParagraphAlignment`, `setParagraphLevel` / `getParagraphLevel`, `setParagraphBullet` / `getParagraphBullet`                                                                                                               |
| Per-run text         | `setShapeRunText` / `getShapeRunText`, `setShapeRunFormat` / `getShapeRunFormat`, `getShapeParagraphCount`, `getShapeRunCount`                                                                                                                          |
| Text frame           | `setShapeTextAnchor` / `getShapeTextAnchor`, `setShapeTextMargins` / `getShapeTextMargins`                                                                                                                                                              |
| Fill                 | `setShapeFill` / `getShapeFill`, `setShapeGradientFill`, `setShapePatternFill`, `setShapeImageFill`, `setShapeNoFill`, `clearShapeFill`                                                                                                                 |
| Stroke               | `setShapeStroke` / `getShapeStroke`, `setShapeStrokeDash` / `getShapeStrokeDash`, `setShapeStrokeArrow` / `getShapeStrokeArrow`, `…NoStroke`                                                                                                            |
| Effects              | `setShapeShadow` / `setShapeGlow` / `getShapeEffect`, `clearShapeEffects`                                                                                                                                                                               |
| Geometry             | `setShapePosition`, `setShapeSize`, `setShapeRotation`, `setShapeFlip`, `setShapeBounds` / `getShapeBounds`                                                                                                                                             |
| Pictures             | `setShapeImage`, `setShapeImageCrop` / `getShapeImageCrop`, `setShapeImageOpacity` / `getShapeImageOpacity`, `setShapeImageBrightness`, `…Contrast`                                                                                                     |
| Z-order              | `bringShapeToFront`, `sendShapeToBack`, `bringShapeForward`, `sendShapeBackward`                                                                                                                                                                        |
| Click actions        | `setShapeClickAction` / `getShapeClickAction` (`url` / `slide` / `nextSlide` / `prevSlide` / `firstSlide` / `lastSlide`)                                                                                                                                |
| Shape removal        | `removeShape`                                                                                                                                                                                                                                           |
| Tables               | `getTableCell` / `getTableCells`, `setTableCellText` / `getTableCellText`, `setTableCellFill` / `clearTableCellFill`, `setTableCellAlignment`, `setTableCellTextFormat`, `insertTableRow` / `removeTableRow`, `insertTableColumn` / `removeTableColumn` |
| Charts               | `addSlideChart`, `getSlideCharts`, `setChartSpec` — kinds: `bar`, `column`, `line`, `pie`, `doughnut`, `area`                                                                                                                                           |
| Theme                | `getPresentationTheme` — color scheme (`accent1`..`accent6`, `dark1`, `light1`, `hyperlink`, ...)                                                                                                                                                       |
| Validation           | `validatePresentation(pres)` — invariant checks, returns `ValidationIssue[]`                                                                                                                                                                            |
| Units                | `inches(n)`, `cm(n)`, `mm(n)`, `pt(n)`, `emu(n)` — return branded `Emu` numbers                                                                                                                                                                         |

## Compatibility

- **Node**: >= 20.
- **Browsers**: current and current-1 of Chrome, Firefox, Safari, Edge.
- **TypeScript**: >= 5.4 (for strict `satisfies` and `const` type parameters).
- **Output**: PPTX files validated against ECMA-376 schemas, smoke-tested
  against PowerPoint (current), Keynote (current), Google Slides, and
  LibreOffice Impress.

## Development

```sh
git clone --recurse-submodules git@github.com:baseballyama/pptx-kit.git
cd pptx-kit
pnpm install
pnpm test
```

If you already cloned without submodules:

```sh
git submodule update --init --recursive --depth 1
```

`references/` holds reference implementations and spec material we read
while building this library. See `references/README.md`.

## Contributing

Before opening an issue or PR, please read `CLAUDE.md` — it documents the
project's design rules, the "one way to do one thing" policy, and what
counts as a real bug report vs. a low-effort AI-generated one.

PRs are expected to:

- Follow the template (`.github/pull_request_template.md`).
- Include a failing test in the same PR that the change makes pass.
- Add a changeset (`pnpm changeset`) for user-visible changes.
- Pass `pnpm typecheck`, `pnpm lint`, and `pnpm test`.

## License

[MIT](./LICENSE)
