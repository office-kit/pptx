---
'pptx-kit': minor
---

feat: `ChartSpec.dataLabels` carries the chart-level `<c:dLbls>` toggles
— `showValue`, `showCategory`, `showSeriesName`, `showPercent` — read
from each plotted-kind element. Playground projects them onto bar /
column tops (numeric value above each bar) and pie / doughnut slices
(value, percent, and / or category text painted at the slice mid-arc).

Adds the `ChartDataLabels` interface to the public type surface.
