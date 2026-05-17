---
'pptx-kit': minor
---

feat: chart `<c:varyColors>` for single-series bar / column.
`ChartSpec.varyColors` carries the `<c:plottedKind><c:varyColors val="1"/>`
flag. When set and the chart has exactly one series, the renderer
assigns each data point a distinct accent color (mirroring
PowerPoint's "Vary colors by point" toggle for column / bar). Pies
already varied colors implicitly.
