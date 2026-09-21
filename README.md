# @office-kit/pptx

[![npm](https://img.shields.io/npm/v/@office-kit/pptx)](https://www.npmjs.com/package/@office-kit/pptx)
[![CI](https://github.com/office-kit/pptx/actions/workflows/ci.yml/badge.svg)](https://github.com/office-kit/pptx/actions/workflows/ci.yml)
[![license](https://img.shields.io/npm/l/@office-kit/pptx)](./LICENSE)

Read, edit, and write `.pptx` (PowerPoint / Office Open XML Presentation) files
from TypeScript, in **Node.js and the browser**, from a single ESM bundle.

**[Create slides with AI](#create-slides-with-ai-and-live-preview)** ·
**[Documentation](https://office-kit.github.io/pptx/)** ·
**[Playground](https://office-kit.github.io/pptx/playground)** (inspect a deck in your browser) ·
**[REPL](https://office-kit.github.io/pptx/repl)** (write code, watch the deck redraw)

```ts
import { loadPresentation, replaceTokensInPresentation, savePresentation } from '@office-kit/pptx';

// A designer makes template.pptx in PowerPoint. Your code fills it in.
const pres = await loadPresentation(templateBytes);
replaceTokensInPresentation(pres, { name: 'Alice', event: 'Re:Invent', date: '2026-12-01' });
const out: Uint8Array = await savePresentation(pres);
```

> **Status: 0.x, pre-1.0.** The capabilities in the [scope table](#scope) are
> exercised against real PPTX fixtures and validated in CI (see
> [How output is checked](#how-output-is-checked)). Until 1.0 the public API is
> not frozen: a breaking change can land in a minor (`0.x`) release, so pin a
> version or an exact range.

## Create slides with AI and live preview

Create presentations in TSX with a live browser preview and Claude Code or Codex
beside the slide. Select an area and ask the agent to change it, edit text
directly, and export an editable PowerPoint file. Changes are saved in the TSX
source; saving a source file also updates the preview.

Requires Node.js 22.18 or later. Create a project and start the preview in one
command (macOS/Linux):

```sh
npx --yes @office-kit/pptx-dev@latest init my-slides && cd my-slides && npm install && npm run dev
```

Open the local URL printed by the server. The generated project includes
`deck.tsx`, individual slides, a shared theme and a `CLAUDE.md` authoring guide.
For later sessions, run this inside the project:

```sh
npx office-pptx dev deck.tsx
```

`npm run dev` starts the same preview. To update an existing project's dev tools,
stop the server, then run:

```sh
npm install -D @office-kit/pptx-dev@latest
npm run dev
```

Run `npm run check` to type-check the deck and `npm run build` to export
`deck.pptx`. AI editing uses your locally installed and authenticated Claude Code
or Codex CLI. After edits, changed-slide screenshots are supplied to the agent
for design review; capture requires Chrome or Playwright Chromium.

See the [preview and agent setup](packages/dev/README.md),
[selection, text editing and visual review](packages/dev/README.md#edit-from-the-slide),
and [TSX element reference](packages/dsl/README.md). To let Claude Code handle
project setup as well, follow the
[skill installation guide](https://office-kit.github.io/pptx/docs/authoring).

## Why this library

- **It reads as well as it writes.** Open a deck made in PowerPoint, Keynote,
  or Google Slides, change it, and save it. Every setter has a getter, and
  there are deck-wide queries (find every hyperlink, every comment by an
  author, every slide with an empty title).
- **Parts it does not model survive the round trip.** SmartArt, OLE objects,
  video, modern threaded comments, and vendor extensions are carried through
  untouched. The library never silently strips what it does not understand.
- **The output is valid, not "valid enough".** Microsoft's own
  `OpenXmlValidator` gates every CI run. A file that opens in PowerPoint but
  breaks Keynote is treated as a bug.
- **One ESM bundle for Node and the browser.** No `fs`, `Buffer`, or `zlib` on
  the hot path, and one runtime dependency ([fflate](https://github.com/101arrowz/fflate), for ZIP).
- **You ship only what you import.** The API uses side-effect-free functions.
  CI checks tree-shaking and bundle-size limits (`test/tree-shake.test.ts`).
- **Types follow the spec.** The model mirrors ECMA-376 Part 1 §19
  (PresentationML). Positions are branded `Emu` numbers, so inches and points
  cannot be mixed up by accident.

## How it differs from PptxGenJS

[PptxGenJS](https://github.com/gitbrent/PptxGenJS) is the established way to
generate a deck in JavaScript, and it is good at that job. The difference is
direction: **PptxGenJS writes new files; `@office-kit/pptx` reads, edits, and
writes.** If your deck starts from a template, or from a file a person made,
PptxGenJS cannot open it.

Compared against PptxGenJS 4.0.1:

|                                   | `@office-kit/pptx`                                                                                              | PptxGenJS                                                                        |
| --------------------------------- | --------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------- |
| Open and edit an existing `.pptx` | ✅ Load, change, save; unknown parts are preserved                                                              | ❌ Creates new files only                                                        |
| Templates                         | Any `.pptx` a designer made in PowerPoint                                                                       | Slide masters defined in code (`defineSlideMaster`)                              |
| Read back what is in a deck       | ✅ Every setter has a getter, plus deck-wide queries                                                            | ❌ The API is write-only                                                         |
| API shape                         | Tree-shakeable functions (`addSlideChart(slide, …)`)                                                            | One class with methods (`slide.addChart(…)`)                                     |
| Module formats                    | ESM only                                                                                                        | ESM, CommonJS, and a script-tag bundle                                           |
| Runtime dependencies              | 1 (`fflate`)                                                                                                    | 4 (`jszip`, `image-size`, `https`, `@types/node`)                                |
| Slide transitions                 | ✅                                                                                                              | ❌                                                                               |
| Animations                        | ✅ Four entrance / exit presets                                                                                 | ❌                                                                               |
| Comments                          | ✅                                                                                                              | ❌ (speaker notes only)                                                          |
| Chart types you can author        | ✅ All 16 ECMA-376 plot types, incl. 3-D, stock, surface, pie-of-pie; combos, error bars, data table, date axis | Bar, line, area, pie, doughnut, scatter, bubble, radar, 3-D bar / bubble; combos |
| Audio, video, YouTube embeds      | ✅ Embedded from bytes (Node and browser), with read-back (`getShapeMedia`); clips de-duplicated per deck       | ✅ From a path or base64; write-only                                             |
| HTML `<table>` to slides          | ❌                                                                                                              | ✅ With automatic paging                                                         |
| Render a slide to an image        | ✅ SVG and PNG, via [`@office-kit/pptx-preview`](packages/preview)                                              | ❌                                                                               |
| How output is checked             | Open XML SDK validator and ECMA-376 XSDs, in CI                                                                 | Manual runs in PowerPoint and other apps before a release                        |

**Pick PptxGenJS** if you only ever generate new decks and need HTML-table
import, CommonJS, or a `<script>`-tag build.

**Pick `@office-kit/pptx`** if a template, an existing deck, or a validation
requirement is involved, or if you need to read a deck as well as write one.

The two also work together. Decks written by PptxGenJS are part of this
library's test fixtures (`test/pptxgenjs-compat.test.ts`), so you can generate
with PptxGenJS and post-process the result here.

### The same slide with PptxGenJS, the core API, and TSX

```ts
// PptxGenJS
import pptxgen from 'pptxgenjs';

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
const slide = pptx.addSlide();
slide.addText('Q3 Review', { x: 1, y: 0.5, w: 8, h: 1, fontSize: 28, bold: true });
slide.addChart(
  pptx.ChartType.bar,
  [{ name: 'Revenue', labels: ['Q1', 'Q2'], values: [120, 180] }],
  {
    x: 1,
    y: 1.5,
    w: 8,
    h: 4,
    barDir: 'col',
  },
);
await pptx.writeFile({ fileName: 'out.pptx' });
```

```ts
// @office-kit/pptx: the Node barrel includes the core API and file helpers.
import * as pptx from '@office-kit/pptx/node';

const { inches } = pptx;
const pres = pptx.createPresentation();
const slide = pptx.addBlankSlide(pres);
const title = pptx.addSlideTextBox(slide, {
  x: inches(1),
  y: inches(0.5),
  w: inches(8),
  h: inches(1),
  text: 'Q3 Review',
});
pptx.setShapeTextFormat(title, { size: 28, bold: true });
pptx.addSlideChart(slide, {
  x: inches(1),
  y: inches(1.5),
  w: inches(8),
  h: inches(4),
  spec: {
    kind: 'column',
    categories: ['Q1', 'Q2'],
    series: [{ name: 'Revenue', values: [120, 180] }],
  },
});
await pptx.savePresentationToFile(pres, 'out.pptx');
```

```tsx
// @office-kit/pptx-dsl — deck.tsx
import { Presentation, Slide, Text, Chart } from '@office-kit/pptx-dsl';

export default (
  <Presentation>
    <Slide>
      <Text x={1} y={0.5} width={8} height={1} size={28} bold>
        Q3 Review
      </Text>
      <Chart
        x={1}
        y={1.5}
        width={8}
        height={4}
        spec={{
          kind: 'column',
          categories: ['Q1', 'Q2'],
          series: [{ name: 'Revenue', values: [120, 180] }],
        }}
      />
    </Slide>
  </Presentation>
);
```

Save the TSX example as `deck.tsx` in an
[initialized authoring project](https://office-kit.github.io/pptx/docs/tsx), then
run `npx --no-install office-pptx build deck.tsx --out out.pptx`.

The core API uses explicit units (`inches(1)`, `cm(2.5)`, `pt(12)`). The DSL uses
inches for numeric geometry and points for font sizes, and compiles to the same
core model. Both produce editable PowerPoint text and charts.

## How output is checked

- **Open XML SDK.** A CI job generates the sample decks and runs Microsoft's
  `OpenXmlValidator` over them (`tools/ooxml-validate`). Any validation error
  fails the build.
- **ECMA-376 XSDs.** Emitted XML is validated against the official schemas
  with `xmllint` in the test suite (these tests skip on a machine without
  `xmllint`).
- **Real files and a parity corpus.** Fixture tests round-trip decks written
  by python-pptx and PptxGenJS, and `test/corpus` authors the same slide once
  with PptxGenJS and once with this library, then diffs the two drawing trees.
  The suite runs on Node 22, 24, and 26.
- **At runtime.** `validatePresentation(pres)` checks package invariants
  (missing relationships, dangling slide ids, duplicate shape ids, layouts
  without masters) in Node and the browser.

## Scope

The work is split into four levels. The `0.x` line covers levels 1–3 and part
of level 4. Items marked "post-1.0" are not implemented yet:

| Level | Capability                                                          | 0.x                             |
| ----- | ------------------------------------------------------------------- | ------------------------------- |
| L1    | Read an existing PPTX, save it back without corruption              | ✅                              |
| L2    | Template edit: text replacement, image swap, add slide from layout  | ✅                              |
| L3    | Authoring: shapes, text, tables, fills, effects, transforms, groups | ✅                              |
| L3    | Authoring on top of existing themes / masters / layouts             | ✅                              |
| L3    | Rebranding a deck: theme colors and theme fonts                     | ✅                              |
| L3    | Constructing new themes / masters / layouts from scratch            | ❌ post-1.0                     |
| L3    | Charts: all 16 plot types (incl. 3-D, stock, surface), and combos   | ✅                              |
| L4    | Notes, comments, transitions                                        | ✅                              |
| L4    | Simple animations (entrance / exit presets)                         | ✅                              |
| L4    | Audio / video / online video: embed, link and read back             | ✅                              |
| L4    | SmartArt authoring                                                  | ❌ post-1.0 (read pass-through) |
| L4    | Complex animation timing trees                                      | ❌ post-1.0                     |
| L4    | OLE / ActiveX authoring                                             | ❌ post-1.0 (read pass-through) |
| L4    | Document encryption (read + write)                                  | ❌ post-1.0                     |

Out-of-scope content is still **preserved on round-trip**. That is the L1
contract.

When NOT to use this:

- You need a **pixel-perfect** PPTX rendering (print, archival). The
  companion [`@office-kit/pptx-preview`](packages/preview) package renders slides to
  SVG in the browser and to PNG on the server, and its closeness to LibreOffice
  is measured per slide and gated in CI (`site/fidelity`). It is a
  high-fidelity preview, not a spec-complete paint engine. For
  pixel-authoritative output, use PowerPoint itself or LibreOffice headless.
- You only generate new decks and need a feature in the PptxGenJS column
  above. Use PptxGenJS.
- You want to convert PPTX to another format (Keynote, ODP). Out of scope
  forever; that is a renderer's job.

## Install

```sh
npm install @office-kit/pptx
# or
pnpm add @office-kit/pptx
# or
yarn add @office-kit/pptx
```

## One API

@office-kit/pptx exposes a single tree-shakeable free-function API. Every
capability is a named export — `loadPresentation`, `savePresentation`,
`addSlideTextBox`, `setShapeFill`, etc. Bundlers drop every entry you
don't import, so the minimal `load → save` bundle is **about 56 KB**
unminified. CI enforces the bound in `test/tree-shake.test.ts`.

```ts
import {
  findSlidePlaceholder,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeText,
} from '@office-kit/pptx';

const pres = await loadPresentation(bytes);
const title = findSlidePlaceholder(getSlides(pres)[0]!, 'title');
if (title) setShapeText(title, 'Hello');
const out = await savePresentation(pres);
```

## Driving @office-kit/pptx from an AI agent

Install the Claude Code plugin once from the shared [Office Kit marketplace](https://github.com/office-kit/skills):

```sh
claude plugin marketplace add office-kit/skills
claude plugin install pptx@office-kit
```

In Claude Code, open `/plugin` → **Marketplaces** → **office-kit** and enable
**auto-update** once. Third-party marketplaces do not enable it by default.
Updates load after `/reload-plugins` or a new session.

Then ask Claude Code:

```text
/pptx:office-kit-pptx Create a quarterly business review with a revenue chart
and a next-actions table. Use labelled sample data where needed.
```

The [skill](skill/SKILL.md) creates a TSX project, starts the interactive preview,
checks the source and exports an editable PPTX. Ask for changes in the same
conversation; the preview updates as the agent edits TSX. You need Node.js 22.18+
and Git alongside Claude Code. See the
[authoring guide](https://office-kit.github.io/pptx/docs/authoring) for template
editing and manual setup. The bundled [core reference](skill/references/core-api.md)
and [tested example](skill/examples/business-deck.md) cover direct API use.

## Usage

### Edit a template

```ts
import {
  findSlidePlaceholder,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeText,
} from '@office-kit/pptx';

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
import { loadPresentation, replaceTokensInPresentation, savePresentation } from '@office-kit/pptx';

const pres = await loadPresentation(templateBytes);
// Replaces `{{name}}`, `{{event}}`, `{{date}}` across every slide.
replaceTokensInPresentation(pres, { name: 'Alice', event: 'Re:Invent', date: '2026-12-01' });
const out = await savePresentation(pres);
```

### Build a deck from scratch (no template file)

`createPresentation()` returns an immediately-authorable deck — a slide
master, the Office theme, and three layouts (`Blank`, `Title Slide`,
`Title and Content`) — with no slides yet. No `.pptx` template needed.

```ts
import {
  addContentSlide,
  addTitleSlide,
  createPresentation,
  findSlideLayoutByType,
  addSlide,
  findSlidePlaceholder,
  savePresentation,
  setShapeText,
} from '@office-kit/pptx';

// Defaults to 16:9; pass { size: '4:3' } for the classic ratio.
const pres = createPresentation();

// Sugar helpers pick the right layout by its locale-stable type token.
addTitleSlide(pres, 'Q3 Business Review');
addContentSlide(pres, { title: 'Agenda', body: 'Highlights and risks' });

// Or bind a layout explicitly. Prefer findSlideLayoutByType — it matches
// the `type` token (`'title'`, `'obj'`, `'blank'`), which is stable
// across PowerPoint UI languages. findSlideLayout(pres, 'Blank') matches
// the user-visible name, which is case-sensitive and localized.
const titleLayout = findSlideLayoutByType(pres, 'title')!;
const slide = addSlide(pres, { layout: titleLayout });
setShapeText(findSlidePlaceholder(slide, 'ctrTitle')!, 'Authored with @office-kit/pptx');

const out: Uint8Array = await savePresentation(pres);
```

### Add slides to your own template

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
} from '@office-kit/pptx';

// Layout names come from the template; list them with getSlideLayouts(pres).
const pres = await loadPresentation(await fetch('/template.pptx').then((r) => r.arrayBuffer()));

const titleLayout = findSlideLayout(pres, 'Title Slide')!;
const slide1 = addSlide(pres, { layout: titleLayout });
setShapeText(findSlidePlaceholder(slide1, 'ctrTitle')!, '@office-kit/pptx demo');
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

### Add video, audio or an online video

```ts
import { addSlideMedia, getShapeMedia, inches } from '@office-kit/pptx';

const box = { x: inches(1), y: inches(1.5), w: inches(8), h: inches(4.5) };

// Embedded video / audio, from bytes. The container is detected from the bytes.
const clip = addSlideMedia(slide, {
  kind: 'video',
  data: mp4Bytes,
  poster: posterPngBytes,
  ...box,
});
addSlideMedia(slide, {
  kind: 'audio',
  data: mp3Bytes,
  x: inches(0.5),
  y: inches(6.5),
  w: inches(0.6),
  h: inches(0.6),
});

// Online video. YouTube watch / youtu.be / shorts URLs become the embed URL.
addSlideMedia(slide2, {
  kind: 'online',
  url: 'https://www.youtube.com/watch?v=dQw4w9WgXcQ',
  ...box,
});

const media = getShapeMedia(clip);
// { kind: 'video', partName: '/ppt/media/media1.mp4', contentType: 'video/mp4', bytes }
// an online video reads back as { kind: 'online', url }
```

The shape is an ordinary picture showing the poster frame, so `setShapeImage`
replaces the poster and `getShapeImageBytes` reads it. Without `poster` a neutral
play-button image is used. Containers detected from the bytes: mp4, m4v, mov,
webm, avi, wmv, mp3, wav, m4a, ogg, wma — pass `format` when the signature
check cannot tell which of those a file is. Identical clip bytes are stored once
per deck, and the clip survives `copyShape`, `duplicateSlide` and `importSlide`.
Whether a clip plays depends on the codecs of the application showing the deck.

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
} from '@office-kit/pptx';

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
import { loadPresentationFile, savePresentationToFile } from '@office-kit/pptx/node';

const pres = await loadPresentationFile('./template.pptx');
await savePresentationToFile(pres, './out.pptx');
```

### Charts

```ts
import {
  addSlideChart,
  getSlides,
  loadPresentation,
  savePresentation,
  inches,
} from '@office-kit/pptx';

const pres = await loadPresentation(templateBytes);
const slide = getSlides(pres)[0];
addSlideChart(slide!, {
  x: inches(0.5),
  y: inches(0.5),
  w: inches(8),
  h: inches(4.5),
  spec: {
    kind: 'column', // bar | column | line | pie | doughnut | area (combine per series with `chartKind`)
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
import { setShapeAnimation, getSlideShapes, getSlides } from '@office-kit/pptx';

const slide = getSlides(pres)[0]!;
const shape = getSlideShapes(slide)[0]!;
setShapeAnimation(shape, { effect: 'fadeIn', durationMs: 800 });
// effects: 'fadeIn' | 'fadeOut' | 'appear' | 'disappear'
```

### Comments

```ts
import { addSlideComment, getSlides } from '@office-kit/pptx';

const slide = getSlides(pres)[0]!;
addSlideComment(slide, {
  author: { name: 'Reviewer A' },
  text: 'Punch up the numbers here.',
  position: { x: 1_000_000, y: 1_000_000 }, // optional EMU coords
});
```

### Gradient fills

```ts
import { setShapeGradientFill } from '@office-kit/pptx';

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
import { validatePresentation } from '@office-kit/pptx';

const issues = validatePresentation(pres);
for (const i of issues) console.error(i.severity, i.message);
// Catches missing rels, dangling slide ids, layouts without masters, etc.
```

### API surface (current state)

Each row lists the free-function entry points. Read/write pairs are
shown together.

| Capability           | API                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| -------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Load / save          | `loadPresentation(input)`, `savePresentation(pres)`, `loadPresentationFile(path)` (node), `savePresentationToFile(pres, path)` (node)                                                                                                                                                                                                                                                                                                                     |
| Create               | `createPresentation({ size?: '16:9' \| '4:3' })` — blank deck with master + theme + `Blank` / `Title Slide` / `Title and Content` layouts                                                                                                                                                                                                                                                                                                                 |
| Slide CRUD           | `getSlides`, `getSlideAt`, `getSlideIndex`, `addSlide`, `removeSlide`, `moveSlide`, `duplicateSlide`, `clearSlideShapes`                                                                                                                                                                                                                                                                                                                                  |
| Slide layout         | `getSlideLayouts`, `findSlideLayout` (by name — case-sensitive, exact; pass a `RegExp` for case-insensitive), `findSlideLayoutByType` (by locale-stable `type` token — preferred), `getSlideLayout(slide)`, `setSlideLayout(slide, layout)`, `getSlideLayoutName`, `getSlideLayoutType`                                                                                                                                                                   |
| Slide metadata       | `getSlideTitle` / `setSlideTitle`, `getSlideSize` / `setSlideSize`, `isSlideHidden` / `setSlideHidden`, `getSlideText`                                                                                                                                                                                                                                                                                                                                    |
| Slide sections       | `getSlideSections`, `setSlideSections` (p14 sectionLst)                                                                                                                                                                                                                                                                                                                                                                                                   |
| Placeholders         | `findSlidePlaceholder(slide, 'title' \| 'body' \| ...)`                                                                                                                                                                                                                                                                                                                                                                                                   |
| Token / text replace | `replaceTokensInPresentation`, `replaceTokensInSlide`, `replaceTextInPresentation`, `replaceTextInSlide`                                                                                                                                                                                                                                                                                                                                                  |
| Background           | `getSlideBackground` / `setSlideBackground` / `clearSlideBackground`                                                                                                                                                                                                                                                                                                                                                                                      |
| Notes                | `getSlideNotes` / `setSlideNotes`                                                                                                                                                                                                                                                                                                                                                                                                                         |
| Transitions          | `getSlideTransition` / `setSlideTransition` / `clearSlideTransition`                                                                                                                                                                                                                                                                                                                                                                                      |
| Animations           | `getShapeAnimation` / `setShapeAnimation` (`fadeIn` / `fadeOut` / `appear` / `disappear`), `clearSlideAnimations`                                                                                                                                                                                                                                                                                                                                         |
| Comments             | `addSlideComment`, `getSlideComments`, `removeSlideComment`, `getCommentAuthors`, `getCommentText` / `getCommentAuthor` / `getCommentPosition`                                                                                                                                                                                                                                                                                                            |
| Shape authoring      | `addSlideTextBox`, `addSlideShape`, `addSlideLine`, `addSlideTable`, `addSlideImage`, `addSlideMedia`, `addSlideChart`                                                                                                                                                                                                                                                                                                                                    |
| Shape lookup         | `findShapeByName`, `findShapesByName`, `findShapesByKind`, `findShapeInPresentation`, `getAllShapes`, `getSlideShapes`                                                                                                                                                                                                                                                                                                                                    |
| Shape text           | `setShapeText`, `setShapeParagraphs`, `setShapeBullets`, `setShapeAlignment`, `setShapeTextFormat`, `setShapeHyperlink` / `getShapeHyperlink`                                                                                                                                                                                                                                                                                                             |
| Per-paragraph        | `setParagraphAlignment` / `getParagraphAlignment`, `setParagraphLevel` / `getParagraphLevel`, `setParagraphBullet` / `getParagraphBullet`, `setParagraphSpacing` / `getParagraphSpacing`, `setParagraphLineSpacing` / `getParagraphLineSpacing`, `getParagraphEndFormat`                                                                                                                                                                                  |
| Per-run text         | `setShapeRunText` / `getShapeRunText`, `setShapeRunFormat` / `getShapeRunFormat`, `getShapeParagraphCount`, `getShapeRunCount`                                                                                                                                                                                                                                                                                                                            |
| Text frame           | `setShapeTextAnchor` / `getShapeTextAnchor`, `setShapeTextMargins` / `getShapeTextMargins`                                                                                                                                                                                                                                                                                                                                                                |
| Fill                 | `setShapeFill` / `getShapeFill`, `setShapeGradientFill`, `setShapePatternFill`, `setShapeImageFill`, `setShapeNoFill`, `clearShapeFill`                                                                                                                                                                                                                                                                                                                   |
| Stroke               | `setShapeStroke` / `getShapeStroke`, `setShapeStrokeDash` / `getShapeStrokeDash`, `setShapeStrokeArrow` / `getShapeStrokeArrow`, `…NoStroke`                                                                                                                                                                                                                                                                                                              |
| Effects              | `setShapeShadow` / `setShapeGlow` / `getShapeEffect`, `clearShapeEffects`                                                                                                                                                                                                                                                                                                                                                                                 |
| Geometry             | `setShapePosition`, `setShapeSize`, `setShapeRotation`, `setShapeFlip`, `setShapeBounds` / `getShapeBounds`                                                                                                                                                                                                                                                                                                                                               |
| Pictures             | `setShapeImage`, `setShapeImageCrop` / `getShapeImageCrop`, `setShapeImageOpacity` / `getShapeImageOpacity`, `setShapeImageBrightness`, `…Contrast`                                                                                                                                                                                                                                                                                                       |
| Z-order              | `bringShapeToFront`, `sendShapeToBack`, `bringShapeForward`, `sendShapeBackward`                                                                                                                                                                                                                                                                                                                                                                          |
| Click actions        | `setShapeClickAction` / `getShapeClickAction` (`url` / `slide` / `nextSlide` / `prevSlide` / `firstSlide` / `lastSlide`)                                                                                                                                                                                                                                                                                                                                  |
| Shape removal        | `removeShape`                                                                                                                                                                                                                                                                                                                                                                                                                                             |
| Tables               | `getTableCell` / `getTableCells`, `setTableCellText` / `getTableCellText`, `setTableCellParagraphs` / `getTableCellParagraphs`, `setTableCellFill` / `clearTableCellFill`, `setTableCellAlignment`, `setTableCellTextFormat`, `insertTableRow` / `removeTableRow`, `insertTableColumn` / `removeTableColumn`, `mergeTableCells` / `getTableCellSpan`                                                                                                      |
| Charts               | `addSlideChart`, `getSlideCharts`, `setChartSpec` — kinds: `bar`, `column`, `line`, `pie`, `doughnut`, `area`; axis tick labels via `categoryAxisTickLabelPos` / `valueAxisTickLabelPos` / `secondaryValueAxis.tickLabelPos`; axis line / gridline widths via `valueAxisLineWidthEmu`, `valueAxisMajorGridlineWidthEmu` and their category / secondary-axis mirrors; series `lineColor` / `markerColor` / `markerLineColor`; `dataLabels.showLeaderLines` |
| Theme                | `getPresentationTheme` / `setPresentationTheme` — color scheme (`accent1`..`accent6`, `dark1`, `light1`, `hyperlink`, ...); `getPresentationFonts` / `setPresentationFonts` — major / minor Latin, East Asian, and complex-script faces                                                                                                                                                                                                                   |
| Groups               | `groupShapes`, `ungroupShapes`, `getGroupChildren`, `getGroupTransform`                                                                                                                                                                                                                                                                                                                                                                                   |
| Text autofit         | `setShapeTextAutoFit` / `getShapeTextAutoFit`, `getShapeTextAutoFitParams`                                                                                                                                                                                                                                                                                                                                                                                |
| Preset adjustments   | `setShapeAdjustValues`; `getShapeCustomGeometry` (custom geometry is read-only)                                                                                                                                                                                                                                                                                                                                                                           |
| Document properties  | `getCoreProperties` / `setCoreProperties`, `touchModified`, `getThumbnail` / `setThumbnail` / `removeThumbnail`                                                                                                                                                                                                                                                                                                                                           |
| Package inspection   | `listPackageParts`, `readPackagePart`, `getMediaParts`, `getOrphanMediaPartNames`, `getPackageSize`, `compactPackage`                                                                                                                                                                                                                                                                                                                                     |
| Validation           | `validatePresentation(pres)` — invariant checks, returns `ValidationIssue[]`                                                                                                                                                                                                                                                                                                                                                                              |
| Units                | `inches(n)`, `cm(n)`, `mm(n)`, `pt(n)`, `emu(n)` — return branded `Emu` numbers                                                                                                                                                                                                                                                                                                                                                                           |

Text formats (`TextFormat`, including a paragraph's `endFormat`) cover the Latin, East Asian and complex-script typefaces (`font`, `fontEastAsian`, `fontComplexScript` — `<a:latin>`, `<a:ea>`, `<a:cs>`). Each is authored and read on its own; setting one leaves the others as they were. Chart labels carry the Latin / East Asian pair and the complex-script slot on `ChartTextStyle`: `font` fills `<a:latin>` and `<a:ea>`, `fontComplexScript` fills `<a:cs>`, and neither implies the other. Rebuilding a typeface writes the `typeface` attribute only, so any `pitchFamily` / `charset` the source file carried on that element is dropped.

Runs that name no face fall back to the theme's font scheme, which is also where a per-script list (`<a:font script="Thai" typeface="Cordia New"/>` and 46 siblings) lives. `createPresentation`'s blank deck carries Office's own list in both `majorFont` and `minorFont`.

Authored XML text and attribute values must contain only XML 1.0 characters. Illegal C0 controls (except tab, LF, and CR), U+FFFE, U+FFFF, and unpaired UTF-16 surrogates throw an error identifying the code point. Remove these characters before authoring; valid supplementary characters such as emoji are preserved.

## Preview and text-overflow checks

[`@office-kit/pptx-preview`](packages/preview) is a companion package that
renders a slide to SVG (browser and Node) or to PNG (Node, via resvg, with no
headless Office). `auditTextLayout` reports text that overflows its box or
wraps unexpectedly, which is how an automated pipeline catches a broken slide
before a person sees it.

```ts
import { getSlides } from '@office-kit/pptx';
import { renderSlideToSvg } from '@office-kit/pptx-preview';

const svg = renderSlideToSvg(pres, getSlides(pres)[0]!);
```

## Office Kit

`@office-kit/pptx` is one of three libraries built on the same rules: the
ECMA-376 spec is the source of truth, output has to validate, and one ESM
build has to run everywhere.

| Package                                                  | Files              |
| -------------------------------------------------------- | ------------------ |
| [`@office-kit/pptx`](https://github.com/office-kit/pptx) | PowerPoint `.pptx` |
| [`@office-kit/xlsx`](https://github.com/office-kit/xlsx) | Excel `.xlsx`      |
| [`@office-kit/docx`](https://github.com/office-kit/docx) | Word `.docx`       |

## Compatibility

- **Node**: >= 22.18 (CI runs 22, 24, and 26).
- **Browsers**: current and current-1 of Chrome, Firefox, Safari, Edge.
- **TypeScript**: >= 5.4 (for strict `satisfies` and `const` type parameters).
- **Output**: validated with the Open XML SDK and the ECMA-376 schemas (see
  [How output is checked](#how-output-is-checked)), and smoke-tested against
  PowerPoint (current), Keynote (current), Google Slides, and LibreOffice
  Impress.

## Development

```sh
git clone --recurse-submodules git@github.com:office-kit/pptx.git
cd pptx
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
