---
'pptx-kit': minor
---

Chart label fonts, table cell merging, and aspect-ratio-preserving image placement

**`ChartTextStyle.font`** — chart titles, axis titles, axis tick labels,
legends, and data labels now accept a font face
(`titleStyle: { font: 'Yu Gothic' }`). The builder writes both
`<a:latin typeface>` and `<a:ea typeface>` so CJK families render
correctly in PowerPoint, and `getSlideCharts` / `getShapeChartSpec` read
the face back for round-trips. Works through both `addSlideChart` and
`setChartSpec`. (The SVG preview keeps its fixed substitution font set —
authored chart faces affect the emitted PPTX, not the preview raster.)

**`mergeTableCells(table, { row, col, rowSpan, colSpan })`** — the write
counterpart to `getTableCellSpan`. Merges a rectangular block into its
top-left anchor, emitting `gridSpan` / `rowSpan` on the anchor and
`hMerge` / `vMerge` on the covered cells per ECMA-376 §21.1.3.18.
Out-of-range blocks, 1×1 blocks, and overlaps with existing merges are
rejected with descriptive errors before anything is mutated.

**`fit: 'contain'` on `addSlideImage` / `setShapeImage`** — preserve the
image's aspect ratio instead of stretching to the target box. `'contain'`
inscribes and centers the image (natural size read from the PNG / JPEG
header); other formats fall back to the default `'fill'` (the existing
stretch behavior) rather than erroring. On `setShapeImage`,
`fit: 'contain'` re-fits the replacement image inside the picture's
current box.
