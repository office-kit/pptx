---
'pptx-kit': minor
---

feat: chart builder writes back per-series `<c:trendline>`. A new
`trendlineElement(tl)` helper emits `<c:trendlineType>`,
`<c:period>` (movingAvg), `<c:order>` (poly), `<c:forward>` /
`<c:backward>`, and `<c:spPr><a:ln><a:solidFill>` color when
authored. Closes the read/write gap for `ChartSeries.trendline`;
round-trip test covers type / period / forward / backward / color.
