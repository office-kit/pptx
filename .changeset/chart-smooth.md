---
'pptx-kit': minor
---

feat(chart): `ChartSeries.smooth` reads `<c:smooth val="1"/>`. Playground
line / area renderer interpolates a cubic-Bézier curve through the
data points (Catmull-Rom-to-Bezier with 0.5 tension) when `smooth` is
true, matching PowerPoint's "smooth line" preset visually.
