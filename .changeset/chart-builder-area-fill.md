---
'pptx-kit': minor
---

feat: chart builder writes back plot-area / chart-area fill + stroke
colors. A new `spPrChildren(fill, stroke)` helper emits
`<c:spPr><a:solidFill><a:srgbClr/></a:solidFill><a:ln>…</a:ln>`. The
builder appends it under `<c:plotArea>` when `plotAreaFill` or
`plotAreaStrokeColor` is set, and under `<c:chartSpace>` (root) when
`chartAreaFill` or `chartAreaStrokeColor` is set. Round-trip test
verifies all four survive.
