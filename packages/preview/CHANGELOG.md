# pptx-kit-preview

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
