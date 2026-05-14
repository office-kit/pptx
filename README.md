# pptx-kit

Generate and edit `.pptx` (PowerPoint / Office Open XML Presentation) files
from TypeScript — in **Node.js or the browser**, from a single ESM bundle.

> **Status: alpha — not yet usable.** The library is in P0 bootstrap. No
> `Presentation` class exists yet. The roadmap and design rationale will be
> published under `docs/` before the first usable release.
>
> Pin exact versions once any 0.x ships.

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

## Usage

> The API is still being built out — examples below describe the intended
> shape and will fill in as features land. See `CHANGELOG.md` for what is
> actually implemented in the current version.

### Author a presentation from scratch

```ts
import { Presentation, inches, pt } from 'pptx-kit';

const pres = Presentation.create({ slideSize: '16:9' });
const slide = pres.slides.add({ layout: pres.slideLayouts.at(1) });
slide.shapes.addText('Hello, OOXML', {
  x: inches(1),
  y: inches(1),
  w: inches(5),
  h: inches(1),
  font: { name: 'Calibri', size: pt(28) },
});

const bytes: Uint8Array = await pres.save();
// Node:    fs.writeFile('out.pptx', bytes)
// Browser: new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' })
```

### Edit a template

```ts
import { Presentation } from 'pptx-kit';

const pres = await Presentation.load(existingPptxBytes);
const cover = pres.slides.at(0);
cover.findPlaceholder('title')?.setText('Q3 Review');

const out = await pres.save();
```

### Node convenience entry

```ts
import { Presentation } from 'pptx-kit/node';

const pres = await Presentation.load('./template.pptx');
await pres.saveTo('./out.pptx');
```

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
