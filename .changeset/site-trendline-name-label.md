---
'pptx-kit-site': patch
---

site(playground): render `ChartTrendline.name` as a small label at the
trendline's right endpoint. Only emitted when the name is authored, so
charts without a custom trendline label look the same as before.
