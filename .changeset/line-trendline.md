---
'pptx-kit': minor
---

feat(chart): line / area charts now overlay the per-series
`<c:trendline>` when authored. Same regression types as the
column-chart variant (linear / log / exp / movingAvg / poly+power
fallback). Only emitted on the clustered layout — stacked plots
already convey the cumulative shape.
