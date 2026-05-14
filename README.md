# pptx-kit

Generate and edit `.pptx` (PowerPoint / Office Open XML Presentation) files
from TypeScript — in **Node.js or the browser**, from a single ESM bundle.

> **Status: pre-1.0, feature-complete for L1–L4 targets.** Every v1.0
> capability in the table below works end-to-end against real PPTX
> fixtures, with every emitted XML part validated against the
> ECMA-376 schemas via `xmllint` in CI. Public API is still subject to
> change — pin exact versions until 1.0.

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
| L3    | Authoring — shapes, text, tables, themes, masters, layouts          | ✅                              |
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

## Two APIs

pptx-kit ships two parallel public surfaces:

1. **Class API** — `Presentation`, `Slide`, `SlideShape`, `SlideLayout`.
   Fluent and discoverable; ideal for quick scripts. Importing
   `Presentation` pulls every authoring method into your bundle.

2. **Tree-shakeable free-function API** — `loadPresentation`,
   `savePresentation`, `addSlideTextBox`, `setShapeFill`, etc. Bundlers
   drop every method you don't import, so the minimal `load → save`
   bundle is **~57 KB** instead of ~171 KB.

Both APIs read and write the same opaque internal state, so values
produced by one work with the other. CI enforces the tree-shake bound
in `test/tree-shake.test.ts`.

```ts
// Class API
import { Presentation } from 'pptx-kit';
const pres = await Presentation.load(bytes);
pres.slides[0]?.findPlaceholder('title')?.setText('Hello');
await pres.save();

// Free-function API — tree-shakeable
import { loadPresentation, findSlidePlaceholder, getSlides, setShapeText, savePresentation } from 'pptx-kit';
const pres2 = await loadPresentation(bytes);
const titleShape = findSlidePlaceholder(getSlides(pres2)[0]!, 'title');
if (titleShape) setShapeText(titleShape, 'Hello');
await savePresentation(pres2);
```

## Usage

### Edit a template

```ts
import { Presentation } from 'pptx-kit';

const pres = await Presentation.load(existingPptxBytes);
const cover = pres.slides[0];
cover.findPlaceholder('title')?.setText('Q3 Review');
cover.findPlaceholder('body')?.setText('Numbers up and to the right.');

const out: Uint8Array = await pres.save();
// Node:    fs.writeFile('out.pptx', out)
// Browser: new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' })
```

### Token-based template fill

```ts
import { Presentation } from 'pptx-kit';

const pres = await Presentation.load(templateBytes);
// Replaces `{{name}}`, `{{event}}`, `{{date}}` across every slide.
pres.replaceTokens({ name: 'Alice', event: 'Re:Invent', date: '2026-12-01' });
const out = await pres.save();
```

### Build a deck from a blank template

```ts
import { Presentation, inches, pt } from 'pptx-kit';

const pres = await Presentation.load(await fetch('/blank.pptx').then((r) => r.arrayBuffer()));

const titleLayout = pres.slideLayouts.find((l) => l.name === 'Title Slide');
const slide1 = pres.addSlide({ layout: titleLayout! });
slide1.findPlaceholder('ctrTitle')?.setText('pptx-kit demo');
slide1.findPlaceholder('subTitle')?.setText('an OOXML library for TypeScript');

const blank = pres.slideLayouts.find((l) => l.name === 'Blank');
const slide2 = pres.addSlide({ layout: blank! });
slide2.addTextBox({
  x: inches(1),
  y: inches(1),
  w: inches(8),
  h: inches(1),
  text: 'Free-form text box',
});
slide2.addImage(imageBytes, { x: inches(1), y: inches(3), w: inches(3), h: inches(3) });

const dup = pres.duplicateSlide(slide2);
pres.moveSlide(dup, 0);

const out: Uint8Array = await pres.save();
```

### Replace an image in place

```ts
import { Presentation } from 'pptx-kit';

const pres = await Presentation.load(templateBytes);
for (const slide of pres.slides) {
  for (const shape of slide.shapes) {
    if (shape.kind === 'picture' && shape.name === 'Logo') {
      shape.setImage(newLogoBytes); // format auto-detected; geometry preserved
    }
  }
}
const out = await pres.save();
```

### Node convenience entry

```ts
import { Presentation } from 'pptx-kit/node';

const pres = await Presentation.loadFile('./template.pptx');
pres.replaceTokens({ name: 'Alice' });
await pres.saveTo('./out.pptx');
```

Or with the free-function path:

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
      { name: 'Cost',    values: [80,  90,  130, 160] },
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
    { offset: 0,   color: '#FF0000' },
    { offset: 1,   color: '#0000FF' },
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

Each row lists the free-function entry point; the class API exposes the
same capability as a method (`pres.save()` vs `savePresentation(pres)`).

| Capability                | API                                                                                                                                      |
| ------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------- |
| Load / save               | `loadPresentation(input)`, `savePresentation(pres)`, `loadPresentationFile(path)` (node), `savePresentationToFile(pres, path)` (node)    |
| Create                    | `createPresentation()`                                                                                                                   |
| Slide CRUD                | `getSlides(pres)`, `addSlide(pres, { layout })`, `removeSlide(pres, slide)`, `moveSlide(pres, slide, toIndex)`, `duplicateSlide(pres, slide)` |
| Slide layouts             | `getSlideLayouts(pres)`, `getSlideLayoutName(layout)`, `getSlideLayoutType(layout)`                                                      |
| Slide size                | `getSlideSize(pres)`, `setSlideSize(pres, opts)`, presets `SLIDE_SIZE_4_3` / `SLIDE_SIZE_16_9` / `SLIDE_SIZE_16_10`                      |
| Slide title               | `getSlideTitle(slide)`, `setSlideTitle(slide, title)`                                                                                    |
| Placeholders              | `findSlidePlaceholder(slide, 'title' \| 'body' \| ...)`                                                                                  |
| Token fill                | `replaceTokensInPresentation(pres, tokens)`, `replaceTokensInSlide(slide, tokens)`                                                       |
| Background                | `setSlideBackground(slide, color)`, `clearSlideBackground(slide)`                                                                        |
| Notes                     | `getSlideNotes(slide)`, `setSlideNotes(slide, text)`                                                                                     |
| Transitions               | `setSlideTransition(slide, opts)`, `clearSlideTransition(slide)`                                                                         |
| Animations                | `setShapeAnimation(shape, { effect: 'fadeIn' \| 'fadeOut' \| 'appear' \| 'disappear' })`, `clearSlideAnimations(slide)`                  |
| Comments                  | `addSlideComment(slide, opts)`, `getSlideComments(slide)`, `removeSlideComment(c)`, `getCommentAuthors(pres)`                            |
| Shape authoring           | `addSlideTextBox`, `addSlideShape`, `addSlideLine`, `addSlideTable`, `addSlideImage`, `addSlideChart`                                    |
| Shape text                | `setShapeText`, `setShapeBullets`, `setShapeAlignment`, `setShapeTextFormat`, `setShapeHyperlink`                                        |
| Per-run text              | `setShapeRunFormat`, `setShapeRunText`, `getShapeParagraphCount`, `getShapeRunCount`, `getShapeRunText`                                  |
| Fill                      | `setShapeFill`, `setShapeGradientFill`, `setShapeNoFill`, `clearShapeFill`                                                               |
| Stroke                    | `setShapeStroke`, `setShapeNoStroke`, `clearShapeStroke`                                                                                 |
| Geometry                  | `setShapePosition`, `setShapeSize`, `setShapeRotation`, `setShapeFlip`; `getShapePosition` / `getShapeSize` / etc.                       |
| Pictures                  | `setShapeImage(bytes)`, `setShapeImageCrop(shape, crop)`                                                                                 |
| Click actions             | `setShapeClickAction(shape, { kind: 'url' \| 'slide' \| 'nextSlide' \| ... })`                                                           |
| Shape removal             | `removeShape(shape)`                                                                                                                     |
| Charts                    | `addSlideChart(slide, { spec: { kind: 'bar' \| 'column' \| 'line' \| 'pie' \| 'doughnut' \| 'area', categories, series } })`             |
| Validation                | `validatePresentation(pres)` — invariant checks, returns `ValidationIssue[]`                                                             |
| Units                     | `inches(n)`, `cm(n)`, `mm(n)`, `pt(n)`, `emu(n)` — return branded `Emu` numbers                                                          |

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
