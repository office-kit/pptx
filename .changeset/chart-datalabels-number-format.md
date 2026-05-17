---
'pptx-kit': minor
---

feat: chart data labels honor `<c:dLbls><c:numFmt formatCode="…"/>`.
`ChartDataLabels.numberFormat` exposes the format code on both
chart-level and per-series toggle groups, and the playground renderer
projects value labels through the same Excel-format subset the value
axis already supports (`"0%"`, `"$#,##0"`, `"0.00"`, etc). Per-series
formats win over the chart-level default.
