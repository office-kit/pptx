---
'pptx-kit': minor
---

feat(chart): `ChartSeries.lineWidthEmu` and `lineDash` read
`<c:ser><c:spPr><a:ln>` per-series stroke width and preset dash.
Playground line / area renderer uses the authored stroke width
(scaled to px) and projects the preset dash to the same
`stroke-dasharray` cadence shape strokes use.
