---
'pptx-kit': minor
---

feat(chart): `ChartSpec.gapWidthPct` and `overlapPct` read from
`<c:gapWidth>` and `<c:overlap>` on bar / column plots. Playground
sizes bars per ECMA-376 §21.2.2.75 — `barW = groupW / (clusterUnits +
gapWidth/100)` with `clusterUnits = 1 + (S - 1)(1 - overlap/100)` —
so authored bar spacing matches PowerPoint instead of the hard-coded
0.8 / 0.7 ratios.
