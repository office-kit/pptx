---
'pptx-kit': minor
---

feat: chart value-axis major gridlines honor authored stroke color.
`ChartSpec.valueAxisMajorGridlineColor` extracts the
`<c:majorGridlines><c:spPr><a:ln><a:solidFill><a:srgbClr/>` color and
the playground renderer paints gridlines with it (falls through to the
existing light-gray default when no color is authored). Branded
templates with custom gridline tints finally render correctly.
