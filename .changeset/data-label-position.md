---
'pptx-kit': minor
---

feat: chart data label position. `ChartDataLabels.position` carries the
`<c:dLbls><c:dLblPos val="…"/>` token (typed as
`ChartDataLabelPosition`). The reader extracts it at both chart-level
and per-series scope. The playground renderer projects `ctr`, `inEnd`,
`outEnd`, `inBase` onto clustered column and bar labels — outside-end
remains the default, but authored positions now move labels inside the
bar or to the base as PowerPoint shows them.
