---
'pptx-kit': minor
---

feat(chart): `ChartAxisScaling.majorUnit` and `minorUnit` read
`<c:valAx><c:majorUnit>` / `<c:minorUnit>` tick spacing. Playground's
value-axis renderer emits ticks at each multiple of the authored
majorUnit instead of nice-rounded auto-ticks when present.
