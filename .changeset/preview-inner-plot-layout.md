---
"@office-kit/pptx-preview": patch
---

fix: chart previews ignored manually positioned inner plot areas (#348)

Plotted data now uses the authored inner plot position and size. Negative sizes collapse to zero, and plot rectangles are clipped to the chart frame so malformed layouts cannot produce negative SVG dimensions or extend the plot over neighboring shapes. Outer layouts retain automatic axis gutters.
