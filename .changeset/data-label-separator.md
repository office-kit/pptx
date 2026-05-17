---
'pptx-kit': minor
---

feat: chart data label separator. `ChartDataLabels.separator` carries
the `<c:dLbls><c:separator>…</c:separator>` text used to join
multiple label parts (value + percent + category etc.). The pie /
doughnut renderer threads the per-series override, falling back to
the chart-level separator and finally to a single space. Common
values: `", "`, `"\n"`, `"; "`.
