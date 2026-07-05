# Plan: template-to-deck authoring & a complete two-runtime preview

Status: **approved** · Owner: maintainers · Last updated: 2026-06-12

This document is the execution plan for @office-kit/pptx's two product goals beyond
"author from scratch" and "load an existing file":

1. **Template-to-deck** — produce a _new_ presentation from an existing
   template or an existing presentation, not just mutate one in place.
2. **Preview** — render any supported presentation faithfully in **both the
   browser and the server (Node)** via `@office-kit/pptx-preview`.

Each workstream below has a concrete deliverable, a test obligation, and an
acceptance criterion. The plan is ordered so that every step is independently
shippable and the measurement infrastructure exists before the fidelity work
it is meant to verify lands.

## 1. Where we are

### Template-to-deck (core library)

The capability is **already reachable** through the public API — this
workstream is verification, not construction:

- Load: `loadPresentation` / `loadPresentationFile` (L1, round-trip safe).
- Mutate: placeholders (`findSlidePlaceholder` + `setShapeText`), token fill
  (`replaceTokensInPresentation`), image swap (`setShapeImageBytes`).
- Recompose: `addSlide` (from a layout), `duplicateSlide`, `importSlide`
  (cross-deck), `mergePresentations`, `removeSlide`, `moveSlide`.

What is missing is a **single end-to-end proof**: one test that starts from a
real template, builds a _new_ deck (keep masters/layouts/theme, drop the
template's slides, add fresh ones), and asserts the output is schema-valid,
opens round-trip clean, and renders.

### Preview (`@office-kit/pptx-preview`)

Two entry points exist today:

- Browser: `renderSlideToSvg` — SVG with `<foreignObject>` text (browser does
  layout).
- Server: `renderSlideToImage` / `renderSlideToRgba` — the same renderer in
  pure-SVG text mode (`textLayout: 'svg'`), rasterized with `@resvg/resvg-js`
  and bundled metric-compatible fonts (Carlito/Caladea/Liberation).

Fidelity is _measured_, not guessed: `site/fidelity/` compares our raster
against LibreOffice-headless PDF exports per slide using **fg-SSIM**
(foreground-weighted SSIM — strict on text ink). Current mean across the 21
generated samples: **≈ 0.66** (LibreOffice 26.2, 1280 px).

Known gaps, from code audit (fallback markers, `data-pptx-preset` stubs) and
the fidelity README:

| Gap                                        | Today's behavior                                                 | Runtime affected |
| ------------------------------------------ | ---------------------------------------------------------------- | ---------------- |
| Vertical text (`vert`, `vert270`, …)       | CSS `writing-mode` in browser; **horizontal** in SVG/server mode | server           |
| Multi-column text (`numCol`)               | CSS columns in browser; **single column** in SVG/server mode     | server           |
| Custom geometry (`a:custGeom`)             | Labelled `<rect>` fallback                                       | both             |
| Table cell text                            | Fixed 18 pt, no per-run format                                   | both             |
| Scatter / radar / bubble charts            | Read as `line` approximation or labelled placeholder             | both             |
| Reflection effect                          | Skipped                                                          | both             |
| Picture bullets (`a:buBlip`)               | "■" fallback                                                     | both             |
| Gradient fill inherited from layout/master | Orange-tint fallback                                             | both             |
| Fidelity CI gate                           | Harness exists, **not in CI**, no committed baseline             | —                |
| Renderer unit tests                        | `render-slide.ts` (5 200 lines) has none                         | —                |

Out of scope for this plan, by design (consistent with the library's own
post-1.0 scope): SmartArt layout, animation/transition playback, 3D effects,
EMF/WMF rasterization. These keep their labelled-fallback behavior, which is
explicit and honest rather than wrong.

## 2. Definition of done

The goal is met when **all** of the following hold:

1. **Template-to-deck**: an E2E test (`test/e2e-template-to-deck.test.ts`)
   builds a new deck from a template fixture, and the output passes schema
   validation, round-trips structurally, and renders via `@office-kit/pptx-preview`
   without fallback markers for supported features.
2. **Runtime parity**: for every supported feature, the SVG/server text path
   produces the same layout decisions as the browser path (same line breaks,
   same column/vertical handling) — asserted by unit tests, not by eye.
3. **Fidelity, measured**: mean fg-SSIM across the sample corpus improves
   from ≈ 0.66 to **≥ 0.80**, with no individual sample regressing below its
   committed baseline. (Table sample `10-tables` is scored against its
   documented LibreOffice divergence, not against PowerPoint behavior.)
4. **Gated**: CI runs the fidelity harness (LibreOffice headless on
   ubuntu-latest) against a committed per-sample baseline and fails on
   regression beyond tolerance.
5. **Tested**: `packages/preview` has its own unit-test surface (SVG
   structure assertions + Node raster smoke tests) executed in CI.
6. **Reviewed**: a full maintainer-perspective code review of the result
   reports zero major findings; CI is green on main.

## 3. Workstreams

Ordered by dependency, then by fidelity impact per unit of risk. Each lands
as its own PR with tests; no PR merges red.

### W1 — SVG-mode text parity (server == browser)

Port vertical text and multi-column layout into the pure-SVG text layouter
(`packages/preview/src/text-layout.ts`):

- `vert` / `vert270` / `eaVert`: lay out into rotated line boxes (transform
  per text block), reusing the existing measurer — no new metrics needed.
- `numCol` + `spcCol`: split the wrapped line list across N column boxes of
  equal width, gap honored.

Tests: layout-level unit tests (line breaks identical to current horizontal
output for the default case; column assignment and rotation transforms for
the new cases). Fidelity samples `vertical-text` and `columns` added to the
corpus.

### W2 — Custom geometry (`a:custGeom`)

The single largest "real template" gap: decks exported from design tools use
`custGeom` heavily.

- Core: add `getShapeCustomGeometry(shape)` exposing the parsed path list
  (`moveTo` / `lnTo` / `arcTo` / `quadBezTo` / `cubicBezTo` / `close`, with
  the `gd` guide-formula evaluator for `avLst`/`gdLst`) from
  `src/internal/drawingml`. Spec: ECMA-376 §20.1.9.
- Preview: translate the path list to an SVG `<path d>` scaled from the
  geometry's `w`/`h` coordinate space to the shape extents, replacing the
  rect fallback. Fill/stroke/effects pipeline unchanged.

Tests: guide-formula evaluator unit tests (the 17 `gd` ops), round-trip
fixture with a real custGeom shape, SVG path assertion, fidelity sample.

### W3 — Table cell text fidelity

Replace the fixed-18 pt cell text with the real resolution chain: cell run
properties → table style text props → theme defaults. Reuses
`getShapeRunFormatEffective`-style cascade already present for body text.

Tests: unit tests per cascade level; fidelity sample with styled table.
The known LibreOffice divergence on built-in table-style banding fills stays
documented and excluded from the gate threshold for that sample.

### W4 — Scatter / radar / bubble charts

- Core (`src/internal/chartml/chart-reader.ts`): stop folding
  `scatterChart` / `radarChart` / `bubbleChart` into `kind: 'line'`; model
  them as their own `ChartKind`s carrying xy(z) tuples (`c:xVal`/`c:yVal`/
  `c:bubbleSize`). Builder support for authoring them is **not** in scope —
  read + render only; authoring stays a documented post-plan item so the
  public-API surface change stays minimal (additive union members only).
- Preview: dedicated plotters — scatter (point markers on two value axes),
  radar (polar polyline), bubble (scatter with area-scaled radii).

Tests: reader fixtures from real PowerPoint files, plot-geometry unit tests,
fidelity samples per kind.

### W5 — Effects & fills polish

- **Reflection** (`a:reflection`): vertically flipped copy with gradient
  opacity mask under the shape.
- **Picture bullets** (`a:buBlip`): surface the blip bytes through the
  existing bullet reader; render as inline image.
- **Inherited gradients**: resolve `a:gradFill` through the shape-style /
  placeholder / layout / master cascade like solid fills already do; delete
  the orange-tint fallback.

Tests: SVG structure assertions for each; fidelity samples.

### W6 — Test infrastructure & CI gate (lands FIRST, before W1)

Build the measurement spine before changing the renderer:

- **Preview unit tests**: a `renderSlideToSvg` assertion helper (parse the
  output, query by `data-pptx-*` attributes) + Node raster smoke test
  (render every generated sample slide; assert non-blank, no thrown errors).
- **Fidelity CI job**: new `fidelity` job in `.github/workflows/ci.yml`
  installing LibreOffice + poppler from apt, pinned via the runner image;
  runs the harness over the generated samples.
- **Committed baseline**: `site/fidelity/baseline.json` — per-sample,
  per-slide fg-SSIM recorded from the gate's own LibreOffice version. Gate
  rule: fail if any slide drops > 0.02 below baseline; a PR that improves
  scores updates the baseline in the same PR. (Absolute targets live in this
  doc, the gate only prevents regressions — LibreOffice version drift moves
  absolute numbers.)
- **Template-to-deck E2E** (Definition-of-done #1).

### W7 — Documentation & release

README capability tables (root + preview) updated to match reality;
changesets per user-visible change (`minor` for new core readers/chart
kinds, `patch`/`minor` for preview); playground caveat text refreshed.

## 4. Execution order & review protocol

```
PR 1  W6  test infra + fidelity CI gate + baseline + E2E
PR 2  W1  SVG text parity
PR 3  W2  custGeom (core reader → preview renderer)
PR 4  W3  table text
PR 5  W4  scatter/radar/bubble
PR 6  W5  reflection / picture bullets / inherited gradients
PR 7  W7  docs + changesets (folded into earlier PRs where natural)
```

Every PR: implementation + tests by a sub-agent where the task is
well-scoped, **always reviewed line-by-line by the lead agent/maintainer
before commit**; quality gates (`format:check`, `lint`, `typecheck`, `test`,
`build`, tree-shake bound) green locally before push; fidelity baseline diff
included in the PR description when scores move.

The plan is complete when §2's six criteria are checked off, the final
full review reports zero major findings, and CI (including the new fidelity
gate) is green on main.
