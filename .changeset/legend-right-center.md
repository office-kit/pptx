---
'pptx-kit': minor
---

feat(site/playground): right / left chart legend stack centers
vertically. Previously the `r` and `l` legend positions both
stacked from a fixed `f.y + 12` top, the same as `tr`. PowerPoint
vertically-centers right / left legends inside the chart area; the
renderer now matches by computing `yStart` from the legend's total
height. `tr` keeps the top-anchored stack.
