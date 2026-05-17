---
'pptx-kit': minor
---

feat: chart value-axis tick labels honor `<c:valAx><c:txPr><a:bodyPr
rot="N"/>`. `ChartSpec.valueAxisLabelRotationDeg` returns the rotation
in degrees (converted from OOXML's 60000ths-of-a-degree). The
playground renders each value-axis tick label with a
`transform=rotate()` around its anchor, symmetric to the
`categoryAxisLabelRotationDeg` we already projected.
