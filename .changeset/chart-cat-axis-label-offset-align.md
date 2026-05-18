---
'pptx-kit': minor
---

feat: `ChartSpec.categoryAxisLabelOffset` and
`ChartSpec.categoryAxisLabelAlign` — two more category-axis tuning
knobs from ECMA-376. `<c:catAx><c:lblOffset val="N"/>` (0..1000, default
100) controls the distance from the axis line to the labels as a
percent of text size; `<c:catAx><c:lblAlgn val="ctr|l|r"/>` controls
how multi-line category labels align relative to their tick mark. Both
are read by chart-reader and written by chart-builder in the correct
CT_CatAx schema order.
