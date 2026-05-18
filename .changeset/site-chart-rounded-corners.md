---
'pptx-kit-site': patch
---

site(playground): honor `ChartSpec.roundedCorners` in the chart-area
backdrop — adds `rx="6" ry="6"` to the chart-area `<rect>` only when
the field is authored true. Mirrors PowerPoint's rounded-corner
treatment without altering the default look of any existing chart.
