---
'pptx-kit': minor
---

feat: `ChartSpec.valueAxis` reports the authored
`<c:valAx><c:scaling>` min / max. Playground respects them when
computing axis ranges, so charts with a fixed authored scale (e.g.
percentage charts pinned to 0..100) render with the same scale the
deck author saw instead of auto-fitting to the data.

Adds the `ChartAxisScaling` interface to the public type surface.
