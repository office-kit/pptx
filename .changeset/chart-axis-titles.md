---
'pptx-kit': minor
---

feat(chart): `ChartSpec.categoryAxisTitle` and `valueAxisTitle` read
the per-axis `<c:title>` rich text on `<c:catAx>` (or `<c:dateAx>` /
`<c:serAx>`) and `<c:valAx>`. Playground paints the value-axis title
rotated -90° along the y-axis and the category-axis title centered
below the x-axis.
