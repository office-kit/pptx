---
'pptx-kit-site': patch
---

site(playground): chart data-label preview honors the new
`ChartDataLabels.textStyle` (sizePt / bold / italic / color). A new
`dataLabelTextAttrs` helper resolves per-series → chart-level cascade
the same way `formatDataLabelValue` does and is threaded into the
bar / column / horizontal-bar / line / area / pie label renderers.
Falls back to the renderer's prior hardcoded size / weight when no
style is authored, so existing previews don't shift.
