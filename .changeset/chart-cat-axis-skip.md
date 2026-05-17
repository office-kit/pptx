---
'pptx-kit': minor
---

feat(chart): category-axis label-skip + position. `ChartSpec.categoryAxisTickLabelSkip`
reads `<c:catAx><c:tickLblSkip val="N"/>` (render every Nth label),
and `categoryAxisTickLabelPos` reads `<c:tickLblPos val="…"/>`
(`'none'` hides labels but keeps the axis line; `'low'`/`'high'`/
`'nextTo'` are the other tokens). Playground honors both — dense
time-series charts with `tickLblSkip="5"` no longer overlap their
labels.
