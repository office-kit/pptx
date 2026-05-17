---
'pptx-kit': minor
---

feat: `getShapeEffects(pres, shape)` returns every effect on the
shape's `<a:effectLst>` (`outerShdw`, `innerShdw`, `glow`, `reflection`,
`softEdge`, `blur`) in document order, with each effect's color
(transform-resolved against the theme), opacity, blur radius, distance,
and angle. PowerPoint composes multiple effects in a single filter
stack — the existing `getShapeEffect` only surfaced the first one.

The playground renderer now emits an SVG `<filter>` chain that
composes the same effects, including a synthesized inner shadow
(SVG has no `feInnerShadow` primitive — built via offset + composite).

Also adds the `ShapeEffectAny` type union to the public surface.
