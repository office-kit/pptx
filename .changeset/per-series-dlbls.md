---
'pptx-kit': minor
---

feat(chart): per-series `<c:ser><c:dLbls>` overrides. `ChartSeries.dataLabels`
mirrors the chart-level `ChartSpec.dataLabels`; the series-level
override wins when present. Playground's bar / column renderers
check the per-series flag first so one series can show labels while
others stay clean — common in financial decks.
