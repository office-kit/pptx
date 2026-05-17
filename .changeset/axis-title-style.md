---
'pptx-kit': minor
---

feat: chart axis titles honor authored `<a:rPr>` font / color.
`ChartSpec.categoryAxisTitleStyle` and `valueAxisTitleStyle` carry the
same `ChartTextStyle` shape as `titleStyle`. The playground renderer
projects size / bold / italic / fill onto both axis title labels,
sharing the helper that drives the chart title.
