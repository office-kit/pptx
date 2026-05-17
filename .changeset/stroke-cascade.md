---
'pptx-kit': minor
---

feat: `getShapeStrokeEffective(pres, shape)` walks the layout → master
placeholder cascade when the shape's own stroke is `'inherit'`. Same
discriminant types (solid / none / inherit) as `getShapeStroke`; first
non-inherit layer wins. Playground uses it so placeholder outlines
authored on the master / layout finally render.
