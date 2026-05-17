---
'pptx-kit': minor
---

feat: pie / doughnut slice explosion. `ChartSeries.pointExplosions`
exposes the per-data-point pull-out percentage from `<c:dPt><c:explosion val="N"/>`,
and the playground renderer offsets exploded slices (and their labels)
outward along the slice mid-angle. Matches the "pulled-out" pie look
authors get from Excel's "Vary colors by point" toggle.
