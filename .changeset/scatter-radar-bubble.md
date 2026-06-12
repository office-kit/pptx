---
'pptx-kit': minor
'@pptx-kit/preview': minor
---

feat: scatter, radar, and bubble charts are now modeled as their own
`ChartKind`s instead of being folded into `line`. `ChartSeries` gains
`xValues` (`<c:xVal>`) and `bubbleSizes` (`<c:bubbleSize>`); `ChartSpec`
gains `scatterStyle`, `radarStyle`, `bubbleScale`, and
`bubbleSizeRepresents`. Read + render only: the preview draws real
scatter (two value axes + markers), radar (polar spokes/rings), and
bubble (area-proportional circles) plots, and the write path now rejects
these kinds loudly — previously a read-modify-write silently corrupted a
scatter chart into a line chart.
