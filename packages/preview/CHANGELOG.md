# pptx-kit-preview

## 0.7.0

### Minor Changes

- 2715289: feat: add `auditTextLayout` — detect text overflowing its box (はみ出し) and unintended soft wraps (段落ち)

  `auditTextLayout(pres, options)` measures every shape's text with the same
  layout engine the preview renders with and reports `overflow-x` / `overflow-y`
  issues (plus opt-in `soft-wrap` reports via `reportSoftWraps`). Results carry
  an `approximate` flag when widths were estimated rather than measured.

  `buildFontkitMeasurer` (the `/node` entry) now accepts `{ fonts }` to register
  the deck's actual font files by their authored family names; registered fonts
  also serve as glyph fallbacks, and glyphs no font covers are estimated
  per-character (CJK ≈ 1em) instead of measured against missing-glyph advances.

## 0.6.2

### Patch Changes

- e09b698: Relax the `engines.node` floor from `>=24.16.0` to `>=22.18.0` on both `pptx-kit` and `pptx-kit-preview` so the maintained LTS lines — Node 22 and Node 24 — are supported, and restore Node 22 to the CI test matrix. The published runtime bundles are unchanged; the previous floor reflected the dev toolchain's pin and needlessly blocked `pnpm install` (under `engine-strict`) on still-supported LTS releases such as Node 22.x and earlier Node 24 LTS patches (e.g. 24.13.x).

## 0.6.1

### Patch Changes

- 4f943a5: Fix several text-on-shape and shape-geometry rendering gaps found by comparing
  output against LibreOffice ground truth on a corpus of realistic, multi-feature
  decks:

  - **Preset pattern fills** (`pct5`–`pct90`, `smGrid`/`lgGrid`, and the
    horizontal/vertical/diagonal hatch families) now match LibreOffice's actual
    substitution for these fills — a density-scaled diagonal hatch for the
    percentage family, and correctly differentiated tile pitches for the "small"
    vs "large" grid variants — instead of a uniformly-dense ordered-dither screen.
  - **Multi-column text bodies** (`numCol` with `noAutofit`/`spAutoFit`) now wrap
    into a new row of columns once the last column also overflows, instead of
    piling all remaining text into the final column forever.
  - **`u="wavy"`** (and `wavyDbl`/`wavyHeavy`) now renders as an actual wavy
    underline in the SVG/raster path (drawn as an explicit path — resvg has no
    `text-decoration-style` support, and the path now scales correctly for a
    superscript/subscript run) and as real CSS in the browser path (without
    also waving a strikethrough on the same run — CSS's `text-decoration-style`
    is a single value for the whole underline + line-through shorthand, and
    PowerPoint always draws strikethrough solid regardless of underline style).
  - **Rotated + vertically-flipped shape text** no longer renders upside-down;
    PowerPoint adds a compensating 180° turn to the text specifically for
    `flip.vertical`, independent of `flip.horizontal`. (A shape nested inside a
    vertically-flipped group is not yet covered by this — see the `KNOWN GAP`
    comment on the group-rendering path in `render-slide.ts`.)
  - **Diagonal connectors that need both a horizontal and vertical flip** (e.g. a
    line drawn from bottom-right to top-left) no longer render reversed — the
    flip was applied twice, once via the endpoint swap and again via a leftover
    transform, cancelling itself out and pointing any arrowhead away from its
    target.
  - **Pie/doughnut data labels** now join as `<category> — <value/percent>`
    (e.g. "Web — 48%"), matching PowerPoint/LibreOffice's order instead of the
    reverse.
  - **Overlapping (non-stacked) area charts** now paint each series' fill AND
    outline back-to-front as one unit, so the first-authored series stays fully
    on top, matching PowerPoint; the category-axis title no longer collides
    with the tick-label row.
  - **Table row/column banding** now uses a pale tint (not a near-solid accent
    color), alternates between two tints across every body row (previously every
    other row was left unshaded), and starts the alternation at the first body
    row rather than one row late.

## 0.6.0

### Minor Changes

- e9eae5c: Sharpen preview-renderer fidelity across the sample corpus. Block arrows now use
  the correct OOXML shaft/head proportions (head length scales with `min(w, h)`),
  text in non-rectangular autoshapes (triangle, diamond, pentagon, star, double
  arrow) wraps inside the shape's inscribed text rectangle, vertical text honours
  the rotated text-box insets, glow effects render as a saturated ring instead of
  a pale haze, hyperlink runs take the theme `hlink` colour, bullets size to the
  paragraph's first run and follow centred/right-aligned text, line breaking is
  space-inclusive (matching LibreOffice/PowerPoint), the first text baseline gets
  the same leading drop for every anchor, and category line charts plot at band
  centres with title/axis text sized in pixels. Overall mean fg-SSIM rises from
  ≈0.82 to ≈0.87.

### Patch Changes

- e9eae5c: Close three renderer correctness gaps surfaced by the expanded corpus:

  - **Preset pattern fills** now render the real ECMA-376 `ST_PresetPatternVal`
    tokens (`horz`/`vert`/`ltHorz`/`ltVert`/`dotGrid` etc.) instead of falling
    through to a 50%-coverage checker — the old matcher keyed on GDI HatchStyle
    names no valid OOXML emits.
  - **`wordArtVert` / `wordArtVertRtl`** stack glyphs upright (one per line) per
    `ST_TextVerticalType`, instead of rotating the run 90°, matching PowerPoint
    and the browser (`text-orientation:upright`) path.
  - **`<a:normAutofit/>` without a baked `fontScale`** now shrinks text to fit the
    box, so overflowing bodies render at the reduced size PowerPoint/LibreOffice
    compute at display time rather than spilling past the box. The shrink factor is
    computed once and shared, so the server (SVG) and browser (`foreignObject`)
    previews agree.

- e9eae5c: Fix a crash when rendering a line or connector that sets an explicit line
  cap or join (`setShapeStrokeCap` / `setShapeStrokeJoin`). The renderer
  emitted both its default `stroke-linecap="round"` and the shape's explicit
  cap on the same element, producing a duplicate SVG attribute that aborted
  the render ("attribute 'stroke-linecap' is already defined"). The default
  is now only applied when the shape does not specify its own.

## 0.5.0

### Minor Changes

- 7200690: Drop Node.js 22 support. The minimum supported version is now Node 24.16. The published runtime bundles are unchanged; this only raises the `engines` floor and the CI/test matrix to Node 24.

## 0.4.0

### Minor Changes

- 3ba1e3d: Drop Node.js 20 support. The minimum supported version is now Node 22.18. The build toolchain moved to tsdown, whose current release requires Node 22.18+; the published runtime bundles are unchanged.

## 0.3.2

### Patch Changes

- 333b19f: fix: chart rendering now matches PowerPoint. The renderer drew an invented
  light-gray chart-area frame, omitted axis spines, used faint inward tick stubs,
  defaulted value-axis gridlines on, rendered every line/scatter marker as a
  circle, and drew bar charts with the category axis upside-down. Now:

  - The chart-area border is drawn only when the chart authors one.
  - Value and category axes draw their spine and outward major tick marks.
  - Major gridlines render only when authored (`<c:majorGridlines>`).
  - Line / scatter / radar markers follow PowerPoint's automatic symbol
    rotation (diamond, square, triangle, x, …) when no symbol is authored.
  - Bar charts order categories bottom-to-top, matching PowerPoint.

- 333b19f: fix: percentage pattern fills (`pct5`…`pct90`) now render at the requested
  coverage. They were drawn as a sparse 1–4 dot grid that read far too light —
  `pct50` looked like ~5% ink instead of a 50% screen. They now use an ordered
  (Bayer) dither so the tone matches PowerPoint.

## 0.3.1

### Patch Changes

- 0f7c538: Preview: take the default text color from the deck's body style, not the `tx1` token

  The preview used `scheme:tx1` as the fallback color for runs without an authored color. On a template with an inverted color map (`tx1 → lt1`) that resolves to the light slot, so body text was painted white on the white background — the whole slide looked blank. PowerPoint instead takes the fallback from the master `bodyStyle` (e.g. `schemeClr bg1`). The preview now does the same via the newly exported `resolveDeckBodyTextColor(slide)`, so default-colored text and table-cell text resolve to the color PowerPoint actually paints.

  - New export **`resolveDeckBodyTextColor(slide)`** — the deck's resolved body-text color (master `bodyStyle`, run through the effective color map + theme). This is the color `addSlideTable` / `addSlideChart` bake in, now reusable by renderers.

## 0.3.0

### Minor Changes

- 4a2ede1: Resolve scheme colors through the slide's color map so inverted-map templates render correctly

  Templates whose slide master inverts the color map (`<p:clrMap bg1="dk1" tx1="lt1">`, common in Google Slides / Canva exports) previously rendered with swapped light/dark colors: slide backgrounds came out black in the preview while PowerPoint paints them white, and generated tables and charts came out with invisible text (the default `tx1` token resolved to the same color as the background).

  - **`getEffectiveColorMap(slide)`** — new export returning the slide's effective color map (the master's `<p:clrMap>` overlaid by a per-slide `<p:clrMapOvr>`). Color resolution and renderers apply it to `schemeClr` tokens before indexing the theme.
  - **`resolveDrawingColor(colorEl, theme, clrMap?)`** — accepts an optional color map; scheme tokens are remapped through it before the theme lookup. Omitting it preserves the previous behavior (correct for the standard map).
  - **`addSlideTable` / `addSlideChart`** now bake the deck's resolved body-text color onto table cells and chart text (axis labels, legend, data labels) so generated tables and charts stay readable regardless of the template's color map. Authored colors still win; override table cells afterwards with `setTableCellTextFormat`.
  - **`pptx-kit-preview`** resolves `schemeClr` tokens through the effective color map, so previews of inverted-map decks match what PowerPoint paints.

## 0.2.0

### Minor Changes

- b03e0cb: feat: fidelity calibration sweep — measured against LibreOffice ground truth,
  mean fg-SSIM rose from ≈0.66 to ≈0.78 (≈0.81 excluding documented
  divergences). Body placeholders now inherit the master `bodyStyle` bullet and
  hanging indent through the paragraph cascade (new `bullet` field on
  `ParagraphProperties` from `getParagraphPropertiesEffective`); charts no
  longer invent a legend when the XML authors no `<c:legend>`, and the value
  axis gets Excel-style headroom above the data max with the tick step
  preserved; the chart builder writes `<c:smooth val="0"/>` explicitly on line
  series (the schema default for an absent element is smooth=1, which made
  LibreOffice draw unauthored lines as curves); and the pure-SVG text layer is
  nudged 0.75px left to land on LibreOffice's pixel grid.

## 0.1.0

### Minor Changes

- e65228f: feat: read custom geometry. New `getShapeCustomGeometry(shape)` returns a
  shape's `<a:custGeom>` (ECMA-376 §20.1.9) as a fully-evaluated path list —
  guide formulas (`avLst`/`gdLst`, all §20.1.9.11 operators) are resolved
  against the shape extents so the returned `moveTo`/`lnTo`/`arcTo`/
  `quadBezTo`/`cubicBezTo`/`close` commands carry only numbers. The preview
  renderer now draws custom geometry as a real SVG path (arcs converted to
  cubic Béziers) instead of a labelled rectangle placeholder; only a custGeom
  that fails to evaluate still falls back, marked `data-pptx-fallback`.
- acc9b15: feat: effects & fills polish. The reflection effect (`a:reflection`) now
  renders as a vertically mirrored, gradient-masked copy honoring start/end
  alpha, distance, and the signed `sy` scale; picture bullets (`a:buBlip`)
  render as real inline images in both text layout modes via the new core
  reader `getParagraphBulletImageBytes` (the "■" fallback remains only when
  bullet bytes are genuinely unavailable); and gradient fills inherited
  through the placeholder layout/master cascade resolve via the new
  `getShapeGradientFillEffective` instead of painting a hardcoded orange
  tint.
- 008e7c1: feat(preview): vertical text (`vert`, `vert270`, `eaVert`, `mongolianVert`,
  `wordArtVert`) and multi-column bodies (`numCol`/`spcCol`) now render in the
  pure-SVG text mode (`textLayout: 'svg'`) used by server-side rasterization.
  Previously they fell back to horizontal single-column layout; now the server
  output matches the browser (`foreignObject`) path: rotated line stacking for
  vertical text and PowerPoint-style sequential column fill for multi-column
  bodies. Browser rendering is unchanged.
- 2207ed1: feat: scatter, radar, and bubble charts are now modeled as their own
  `ChartKind`s instead of being folded into `line`. `ChartSeries` gains
  `xValues` (`<c:xVal>`) and `bubbleSizes` (`<c:bubbleSize>`); `ChartSpec`
  gains `scatterStyle`, `radarStyle`, `bubbleScale`, and
  `bubbleSizeRepresents`. Read + render only: the preview draws real
  scatter (two value axes + markers), radar (polar spokes/rings), and
  bubble (area-proportional circles) plots, and the write path now rejects
  these kinds loudly — previously a read-modify-write silently corrupted a
  scatter chart into a line chart.

### Patch Changes

- ba94f5e: fix(preview): table cell text now renders with its real per-run formatting —
  font size, bold, italic, color, typeface, and paragraph alignment — and wraps
  within the cell width, in both the browser (`foreignObject`) and server
  (`svg`) text modes. Previously every cell was drawn at a flat 18 pt with no
  styling. Cells with no explicit run size still fall back to PowerPoint's 18 pt
  default in the theme's body font.
