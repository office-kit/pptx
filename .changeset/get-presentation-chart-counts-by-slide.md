---
'pptx-kit': minor
---

feat: `getPresentationChartCountsBySlide(pres)` — dense per-slide chart
count array. Counts every chart returned by `getSlideCharts` regardless
of whether its spec parsed; pair with `getPresentationChartKindCounts`
for kind-level totals. Rounds out the density-array family alongside
`getPresentationCommentCountsBySlide`,
`getPresentationShapeCountsBySlide`, and
`getPresentationTextLengthsBySlide`.
