---
'pptx-kit': minor
---

feat: `setShapeTextBodyRotationDeg(shape, rotationDeg | null)` — companion
writer for the existing `getShapeTextBodyRotationDeg` reader. Sets
`<a:bodyPr rot="N"/>` (in 60000ths of a degree, per OOXML) so the text
body can rotate independently of the shape's own `<p:xfrm rot>`. Passing
`null` or `0` clears the attribute so the shape inherits the default.
