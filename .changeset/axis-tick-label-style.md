---
'pptx-kit': minor
---

feat: chart axis *tick labels* honor authored `<c:txPr>` font / color.
`ChartSpec.categoryAxisLabelStyle` and `valueAxisLabelStyle` carry the
font / color extracted from `<c:catAx><c:txPr>` and `<c:valAx><c:txPr>`.
A shared `axisTickAttrs` helper composes the SVG `font-*` / `fill`
attributes; the value-axis renderer and category-axis renderer both
project it onto every tick label.
