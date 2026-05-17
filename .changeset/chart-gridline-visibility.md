---
'pptx-kit': minor
---

feat(chart): `ChartSpec.valueAxisMajorGridlines` / `valueAxisMinorGridlines`
read the presence of `<c:majorGridlines/>` / `<c:minorGridlines/>`
under `<c:valAx>`. Playground hides gridlines when `majorGridlines`
is explicitly `false` (absent in the source XML) — common on KPI
charts that show clean bars / lines without horizontal rules behind
them. Tick labels still render.
