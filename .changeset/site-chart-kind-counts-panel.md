---
'pptx-kit-site': patch
---

site(playground): surface chart-kind counts in the summary panel.
When the deck has any charts, a new "chart kinds" row shows the
counts per kind (e.g. `3 column · 2 pie · 1 line`). Built on the
`getPresentationChartKindCounts` aggregator.
