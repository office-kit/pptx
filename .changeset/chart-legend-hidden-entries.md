---
'pptx-kit': minor
---

feat: chart `<c:legend><c:legendEntry><c:delete val="1"/>` honored.
`ChartSpec.legend.hiddenIndices` carries the series indices the
author wants suppressed from the legend (typically trendline series).
The playground filters the parallel legend arrays (names, colors,
marker glyphs) in lock-step so the remaining entries stay aligned,
without affecting plotted data.
