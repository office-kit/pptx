---
'pptx-kit': minor
---

feat(chart): `ChartSeries.invertIfNegative` reads `<c:ser>
<c:invertIfNegative val="1"/>`. Playground's bar / column renderer
paints negative bars in a darker shade of the series color when the
flag is set — matching PowerPoint's profit/loss visualization.
