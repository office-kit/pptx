---
'pptx-kit': minor
---

feat: `ChartSeries.pointColors` — sparse map of per-data-point fill
overrides read from `<c:ser><c:dPt><c:spPr><a:solidFill>`. Pie /
doughnut decks almost always emit one of these per slice; the playground
now paints each slice in its authored color (and reflects it in the
legend swatches) rather than cycling through the accent palette.
