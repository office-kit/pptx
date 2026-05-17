---
'pptx-kit': minor
---

feat: chart builder writes back value-axis extras and tick marks.
`<c:valAx>` now emits `<c:scaling><c:logBase>`, `<c:majorTickMark>`,
and `<c:dispUnits><c:builtInUnit>` when authored on `ChartSpec`;
`<c:catAx>` emits `<c:majorTickMark>`. Round-tripping a chart with
these fields no longer drops them. Covered by a new round-trip test
in `fn-chart-readback`.
