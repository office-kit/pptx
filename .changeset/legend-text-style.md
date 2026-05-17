---
'pptx-kit': minor
---

feat: chart legend honors authored `<c:txPr>` font / color.
`ChartSpec.legend.textStyle` carries the same `ChartTextStyle` shape
used for the chart title and axis titles. The playground renderer
projects font-size, bold, italic, and fill color onto every legend
label across all four position layouts (right / left / top / bottom /
top-right).
