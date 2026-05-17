---
'pptx-kit': minor
---

feat: `getShapeTextBodyRotationDeg(shape)` returns the shape's text-body
rotation from `<a:bodyPr rot="N"/>` (where N is in 60000ths of a
degree). Distinct from the shape's geometry rotation (`<p:xfrm rot>`):
this rotates the text body *inside* the shape without rotating the
geometry. The playground renderer pivots the text body around the
inset midpoint when the angle is non-zero, matching PowerPoint's
behaviour for vertical-label callouts and rotated text frames.
