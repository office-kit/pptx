---
'pptx-kit': minor
---

feat: chart-area / plot-area authored outline strokes.
`ChartSpec.chartAreaStrokeColor` reads `<c:chartSpace><c:spPr><a:ln>`;
`ChartSpec.plotAreaStrokeColor` reads `<c:plotArea><c:spPr><a:ln>`.
The playground renderer projects them onto the chart-area card
border and the plot-area inner rect — branded charts with thick / no
/ colored card borders finally render the way PowerPoint shows them.
