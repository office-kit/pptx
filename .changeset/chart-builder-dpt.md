---
'pptx-kit': minor
---

feat: chart builder writes back per-data-point `<c:dPt>` overrides.
New `dPtElements(colors, explosions)` helper emits sparse
`<c:dPt><c:idx><c:bubble3D val="0"/>[<c:explosion>]
[<c:spPr><a:solidFill><a:srgbClr/>]</c:dPt>` entries from
`ChartSeries.pointColors` and `ChartSeries.pointExplosions`.
Round-trip test asserts both sparse arrays survive.
