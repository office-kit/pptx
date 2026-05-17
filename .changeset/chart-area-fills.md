---
'pptx-kit': minor
---

feat(chart): `ChartSpec.chartAreaFill` and `plotAreaFill` read
`<c:chartSpace><c:spPr><a:solidFill>` and `<c:plotArea><c:spPr>
<a:solidFill>`. Playground paints the chart-area backdrop in the
authored color (replacing the hard-coded white) and adds a tinted
rect behind the plot area when `plotAreaFill` is authored. Common
on branded dashboards that paint a subtle background behind the
chart bars.
