---
'pptx-kit': minor
---

feat(chart): `ChartSpec.valueAxisHidden` and `categoryAxisHidden`
read `<c:valAx><c:delete val="1"/>` and `<c:catAx><c:delete val="1"/>`.
Playground skips rendering the axis when hidden — common on KPI tile
charts that show just the data points without axis labels.
