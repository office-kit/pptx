---
'pptx-kit-site': patch
---

site(playground): project `ChartSpec.valueAxisTitleRotationDeg` and
`ChartSpec.categoryAxisTitleRotationDeg` into the chart axis-title
SVG. Authored rotation overrides the renderer's defaults (`-90` for
value axis, `0` for category axis). Pivot stays at the label anchor
so the title hugs its axis.
