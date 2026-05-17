---
'pptx-kit': minor
---

feat(chart): per-series marker symbol + size.
`ChartSeries.markerSymbol` / `markerSizePt` read `<c:ser><c:marker>`
(`<c:symbol val="…"/>` + `<c:size val="N"/>`). Playground line / area
renderer emits the matching SVG glyph at each data point — circle /
square / diamond / triangle / star / x / plus / dash / dot — sized
per the authored point value. `none` hides the markers.
