---
'pptx-kit': minor
---

feat: `ChartSpec.categoryAxisLineColor` and `valueAxisLineColor` —
authored stroke color on the axis line itself
(`<c:catAx|valAx><c:spPr><a:ln><a:solidFill><a:srgbClr val=…/>`).
`undefined` falls back to the renderer's default. Read by chart-reader,
written by chart-builder in the correct CT_CatAx / CT_ValAx schema
order (after the tick-mark elements, before `<c:txPr>`).
