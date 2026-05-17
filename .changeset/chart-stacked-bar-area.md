---
'pptx-kit': minor
---

feat(site/playground): bar (horizontal), line, and area charts now
honour `ChartSpec.grouping` for stacked / percentStacked layouts —
matching the column-chart treatment added previously. Data labels
render inside the stacked segments for bar (white bold), at the
appropriate cumulative position for line / area, and percent-stacked
versions normalize each category to 100%.
