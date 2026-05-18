---
'pptx-kit': minor
---

feat: `ChartSpec.categoryAxisMajorGridlines` and
`ChartSpec.categoryAxisMinorGridlines` — companions to the existing
`valueAxis*` pair. Bar charts (where the category axis sits on the
vertical edge) actually use these as horizontal guide lines per
category band. Mapped to `<c:catAx><c:majorGridlines/>` /
`<c:minorGridlines/>`. Read by chart-reader, written by chart-builder
in the correct CT_CatAx schema order (right after `<c:axPos>`).
