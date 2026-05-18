---
'pptx-kit': minor
---

feat: `ChartSpec.categoryAxisTickMarkSkip` — the second half of the
ECMA-376 `<c:catAx>` skip pair. `<c:tickLblSkip>` (already supported)
controls label-skip stride; `<c:tickMarkSkip val="N"/>` independently
draws every Nth tick mark. Useful when you want fewer label collisions
but the same dense tick lattice. Read by chart-reader and written by
chart-builder.
