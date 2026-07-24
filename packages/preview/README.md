# @office-kit/pptx-preview

Preview renderer for [`@office-kit/pptx`](https://github.com/office-kit/pptx).
Turns a `@office-kit/pptx` slide model into an **SVG** (browser + Node) or rasterizes
it to a **PNG / RGBA image in Node — with no headless browser**.

> **Experimental (0.x).** This package lives in the `@office-kit/pptx` monorepo and
> also powers the docs-site playground and the fidelity harness. The renderer
> is an _approximation_ of PowerPoint / LibreOffice output and is still
> evolving; the API may change between minor versions, and `0.x` semver applies
> (a minor bump may break). See [fidelity](#fidelity) below.

## Why

`@office-kit/pptx` core does not render — by design. But "show me this deck" comes up
constantly: a docs playground, a thumbnail service, a visual-diff test. The
hard requirement is that rendering must work **in Node and rasterize to an
image without spawning a browser**, so it fits CI and serverless. This package
lays text out as pure SVG `<text>` (no `<foreignObject>`) and paints it with
[resvg](https://github.com/yisibl/resvg-js), which has no browser dependency.

## Entry points

| Import                          | Runtime        | Use                                                 |
| ------------------------------- | -------------- | --------------------------------------------------- |
| `@office-kit/pptx-preview`      | browser + Node | `renderSlideToSvg` → an SVG string                  |
| `@office-kit/pptx-preview/node` | Node only      | `renderSlideToImage` / `renderSlideToRgba` → pixels |

The browser entry pulls in **no** Node built-ins (no `node:fs`, resvg, or
fontkit), so it bundles cleanly for the web.

## Usage

### SVG (browser or Node)

```ts
import { renderSlideToSvg } from '@office-kit/pptx-preview';
import { loadPresentation, getSlides } from '@office-kit/pptx';

const pres = await loadPresentation(bytes);
const svg = renderSlideToSvg(pres, getSlides(pres)[0]);
// → '<svg …>…</svg>'  (text laid out via <foreignObject> — the browser wraps it)
```

### PNG / RGBA (Node, no browser)

```ts
import { renderSlideToImage, renderSlideToRgba } from '@office-kit/pptx-preview/node';
import { loadPresentationFile, getSlides } from '@office-kit/pptx/node';

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

### Text-layout audit (overflow / soft-wrap detection)

```ts
import { auditTextLayout } from '@office-kit/pptx-preview';
import { buildFontkitMeasurer } from '@office-kit/pptx-preview/node';

// Register the fonts the deck actually uses (especially CJK) for accurate widths.
const measureText = buildFontkitMeasurer({
  fonts: [{ family: '游ゴシック', source: 'fonts/YuGothic.ttf' }],
});

const issues = auditTextLayout(pres, { measureText, reportSoftWraps: true });
// → [{ kind: 'overflow-y', slideIndex: 2, shapeName: 'Title 1', overflowPx: 14.3, approximate: false }, …]
```

`auditTextLayout` measures every shape's text body — placeholders, text boxes,
autoshapes, shapes inside groups — with the same engine the preview renders
with, honouring the effective bodyPr cascade (insets, wrap, anchor, vertical
text, columns) and `normAutofit` shrinking. It returns one issue per finding:

| `kind`       | Meaning                                                                       | Extra fields                   |
| ------------ | ----------------------------------------------------------------------------- | ------------------------------ |
| `overflow-x` | Text ink escapes the box horizontally (typically `wrap="none"`)               | `overflowPx`                   |
| `overflow-y` | Wrapped text is taller than the box (or escapes the top when bottom-anchored) | `overflowPx`                   |
| `soft-wrap`  | A paragraph wrapped onto more lines than its authored `<a:br>` count (段落ち) | `paragraphIndex`, `extraLines` |

Every issue also carries `slideIndex`, `shapeName`, and `approximate` —
`true` when a width had to be estimated (the heuristic measurer, or glyphs the
available fonts don't cover), in which case treat borderline verdicts as
advisory.

Options:

- `measureText` — the measurer. Defaults to the browser-safe heuristic
  (≈0.55 em per Latin glyph, 1 em per CJK glyph, always `approximate`); pass
  `buildFontkitMeasurer()` in Node for real glyph metrics.
- `tolerancePx` — overflow at or below this many px (96 DPI) is ignored.
  Default `1`; raise it if you only care about visibly broken slides.
- `reportSoftWraps` — opt into `soft-wrap` issues. Off by default because
  wrapping is normal for body text; turn it on when auditing titles or labels
  meant to stay on one line.

Accuracy: the bundled fonts are metric-compatible with the Office defaults, so
verdicts for Calibri / Cambria / Arial / Times / Courier decks match real glyph
widths; for any other font (custom brand fonts, Japanese fonts), register the
real files via `buildFontkitMeasurer({ fonts })` — a registered font is used
both for runs that name it and as a glyph fallback for CJK text in runs that
resolve to a Latin face. Line-break positions can differ from PowerPoint by a
few characters in edge cases (kinsoku, hyphenation), which is what the default
1 px tolerance absorbs. Table cell text is not audited yet.

## Fidelity

This is a high-fidelity preview, not a spec-complete PowerPoint renderer.
Preset and custom geometry, solid/gradient/pattern/image fills (including the
placeholder layout/master cascade), strokes, rotation, effects (shadow, glow,
soft edge, reflection), images with adjustments, charts (column, bar, line,
area, pie, doughnut, scatter, radar, bubble), tables with per-run cell text,
vertical and multi-column text in both text-layout modes, picture bullets, and
template (layout/master) decoration all render. SmartArt, animations, 3D, and
EMF/WMF fall back to labelled placeholders carrying a machine-readable marker
(below). Per-slide closeness to a LibreOffice baseline is measured and gated
in CI by the fidelity harness in the monorepo (`site/fidelity`) — mean
fg-SSIM ≈ 0.78 across the corpus, with the residual gaps documented there.

### Fallback markers

When a shape cannot be rendered (unsupported format, missing bytes, or
unrecognised content type), the renderer emits a labelled placeholder rectangle.
The placeholder's top-level `<g>` element carries a `data-pptx-fallback`
attribute so automated tooling can detect partial renders without string-parsing
the label text:

| Value            | Trigger                                                     |
| ---------------- | ----------------------------------------------------------- |
| `"image"`        | Image bytes missing (external link) or format not decodable |
| `"chart"`        | Chart kind not modelled by this renderer                    |
| `"graphicFrame"` | Graphic frame with no recognised content (SmartArt, etc.)   |
| `"custGeom"`     | Shape uses custom geometry (`<a:custGeom>`)                 |

Example: `svg.querySelectorAll('[data-pptx-fallback]')` lists every shape that
did not fully render.

## License

MIT (code). Bundled fonts: OFL-1.1 / Apache-2.0 — see `fonts/LICENSES.md`.
