---
'pptx-kit-site': patch
---

site(playground): project `ChartSpec.valueAxisLineColor` and
`ChartSpec.categoryAxisLineColor` into the chart preview. Draws an
explicit axis spine at the appropriate edge (left for vertical / bottom
for horizontal, swapped for bar charts) only when the line color is
authored, so unchanged charts look the same.
