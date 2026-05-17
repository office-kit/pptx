---
'pptx-kit': minor
---

feat(chart): `ChartSpec.titleOverlay` and `ChartSpec.legend.overlay`
read `<c:title><c:overlay>` / `<c:legend><c:overlay>`. When `true`,
the title / legend sits on top of the plot area instead of taking a
horizontal strip. Playground sizes the plot area accordingly — gives
the chart back the extra vertical real estate when overlay is set.
