---
'pptx-kit': minor
---

feat: `ChartTrendline.displayEquation` and `ChartTrendline.displayRSquared`
— two booleans that toggle the regression-equation label and R²
coefficient overlay next to a trendline. Map to
`<c:trendline><c:dispEq val="1"/>` and `<c:dispRSqr val="1"/>`. Read by
chart-reader; written by chart-builder in the correct CT_Trendline
schema order (after `<c:backward>`, before any `<c:trendlineLbl>`).
