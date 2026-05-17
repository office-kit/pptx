---
'pptx-kit': minor
---

feat(chart): `ChartSeries.trendline` reads `<c:trendline>` —
regression type (linear / exp / log / poly / power / movingAvg),
moving-average period, polynomial order, and the trendline's stroke
color. Playground overlays a dashed trendline on bar / column / line
charts; linear / log / exp use fitted regressions, movingAvg uses a
rolling mean.

Adds the `ChartTrendline` type to the public surface.
