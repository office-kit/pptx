---
'pptx-kit': minor
---

feat: `ChartSpec.roundedCorners` — round-trip the chartSpace-level
`<c:roundedCorners val>` toggle. PowerPoint's default is `false`; the
reader surfaces `true` only when the wire is explicitly `1` and the
builder emits the element only when authored, so common defaults stay
clean. Schema position is BEFORE `<c:chart>` (per CT_ChartSpace).
