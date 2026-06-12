# pptx-kit-preview

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
