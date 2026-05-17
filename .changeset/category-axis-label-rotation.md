---
'pptx-kit': minor
---

feat: chart category-axis tick labels honor `<a:bodyPr rot="N"/>`.
`ChartSpec.categoryAxisLabelRotationDeg` carries the authored rotation
(converted from OOXML's 60000ths-of-a-degree to plain degrees). The
playground renderer rotates each tick label around its anchor and
shifts the text-anchor side based on the sign of the rotation so dense
charts with 45°/-45°/90° rotated labels render the way PowerPoint
shows them. Rotated labels also get a longer truncation budget before
ellipsization.
