---
'pptx-kit': minor
---

feat: `findSlidesWithChartKind(pres, kind)` — kind-filtered variant of
the existing `getSlidesWithCharts`. Returns every slide carrying at
least one chart of the given `ChartKind` (`'bar'`, `'column'`,
`'line'`, `'pie'`, `'doughnut'`, `'area'`). Built on `getSlideCharts`
so the predicate respects the spec the renderers actually see.
