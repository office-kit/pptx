# @pptx-kit/preview

Preview renderer for [`pptx-kit`](https://github.com/baseballyama/pptx-kit).
Turns a `pptx-kit` slide model into an **SVG** (browser + Node) or rasterizes
it to a **PNG / RGBA image in Node — with no headless browser**.

> **Experimental (0.x).** This package lives in the `pptx-kit` monorepo and
> also powers the docs-site playground and the fidelity harness. The renderer
> is an _approximation_ of PowerPoint / LibreOffice output and is still
> evolving; the API may change between minor versions, and `0.x` semver applies
> (a minor bump may break). See [fidelity](#fidelity) below.

## Why

`pptx-kit` core does not render — by design. But "show me this deck" comes up
constantly: a docs playground, a thumbnail service, a visual-diff test. The
hard requirement is that rendering must work **in Node and rasterize to an
image without spawning a browser**, so it fits CI and serverless. This package
lays text out as pure SVG `<text>` (no `<foreignObject>`) and paints it with
[resvg](https://github.com/yisibl/resvg-js), which has no browser dependency.

## Entry points

| Import                   | Runtime        | Use                                                 |
| ------------------------ | -------------- | --------------------------------------------------- |
| `@pptx-kit/preview`      | browser + Node | `renderSlideToSvg` → an SVG string                  |
| `@pptx-kit/preview/node` | Node only      | `renderSlideToImage` / `renderSlideToRgba` → pixels |

The browser entry pulls in **no** Node built-ins (no `node:fs`, resvg, or
fontkit), so it bundles cleanly for the web.

## Usage

### SVG (browser or Node)

```ts
import { renderSlideToSvg } from '@pptx-kit/preview';
import { loadPresentation, getSlides } from 'pptx-kit';

const pres = await loadPresentation(bytes);
const svg = renderSlideToSvg(pres, getSlides(pres)[0]);
// → '<svg …>…</svg>'  (text laid out via <foreignObject> — the browser wraps it)
```

### PNG / RGBA (Node, no browser)

```ts
import { renderSlideToImage, renderSlideToRgba } from '@pptx-kit/preview/node';
import { loadPresentationFile, getSlides } from 'pptx-kit/node';

const pres = await loadPresentationFile('deck.pptx');
const slide = getSlides(pres)[0];

// PNG-encoded bytes:
const png = renderSlideToImage(pres, slide, { width: 1280 });

// Raw RGBA pixels (+ the same frame PNG-encoded), for SSIM / diffing:
const { image, png: png2 } = renderSlideToRgba(pres, slide, { width: 1280 });
// image: { width, height, data: Uint8Array }  // row-major RGBA
```

The Node path lays text out as pure `<text>` and measures it with a fontkit
measurer over **bundled** metric-compatible fonts (Carlito ≈ Calibri, Caladea ≈
Cambria, Liberation ≈ Arial/Times/Courier; OFL / Apache-2.0, see
`fonts/LICENSES.md`). The measurer, resvg's font set, and the SVG family names
all reference the same fonts, so wrap/positioning math agrees with the painted
glyphs and the result is deterministic (no system fonts).

## Fidelity

This is an approximation, not a spec-complete PowerPoint renderer. Preset
geometry, fills, strokes, rotation, images, charts, tables, and template
(layout/master) decoration render; custom geometry, SmartArt, and effects are
partial or fall back to labelled placeholders. Per-slide closeness to a
LibreOffice baseline is tracked by the fidelity harness in the monorepo
(`site/fidelity`).

## License

MIT (code). Bundled fonts: OFL-1.1 / Apache-2.0 — see `fonts/LICENSES.md`.
