---
'pptx-kit': minor
---

feat: `ChartSpec.plotVisibleCellsOnly` — toggle `<c:plotVisOnly val/>`.
PowerPoint's default is `true` (only plot visible cells); the field
exists to let authors opt into `false` (plot hidden rows / columns too).
The reader surfaces `false` only when the wire is explicitly `0` so
round-tripping the common default doesn't drag a redundant explicit
`true` into the spec.
