---
'pptx-kit': minor
---

feat(chart): `ChartSpec.firstSliceAngleDeg` reads `<c:firstSliceAng>`
and `ChartSpec.holeSizePct` reads `<c:holeSize>` for doughnut charts.
Playground rotates the first slice's starting position clockwise from
12 o'clock per the authored angle, and sizes the doughnut hole at the
authored percent (10..90) of the outer radius instead of the
hard-coded 55%.
