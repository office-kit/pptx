---
'pptx-kit': minor
---

feat(chart): `ChartSpec.dropLines` and `hiLowLines` read
`<c:dropLines>` and `<c:hiLowLines>` on line / area / stock plots.
Playground renders drop lines from each first-series data point down
to the value axis (dashed gray) and hi-low lines as a vertical span
between the highest and lowest series value at each category
(solid darker gray). The latter is the canonical OHLC pattern.
