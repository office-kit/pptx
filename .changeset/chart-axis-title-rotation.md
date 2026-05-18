---
'pptx-kit': minor
---

feat: `ChartSpec.categoryAxisTitleRotationDeg` and
`ChartSpec.valueAxisTitleRotationDeg` — rotation in plain degrees
(clockwise) on the per-axis title. Maps to
`<c:catAx|valAx><c:title><c:tx><c:rich><a:bodyPr rot="N"/>` (60000ths
of a degree on the wire). PowerPoint often emits `-90` on the value-
axis title; the field now survives round-trip. Read by chart-reader
via a new `readTitleRotationDeg` helper; written by chart-builder
through an extended `titleElement(title, style?, rotationDeg?)`
signature.
