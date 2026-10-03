---
"@office-kit/pptx": minor
---

`setShapeShadow` and `setShapeGlow` no longer erase each other. Each now
replaces only the effect of its own kind and leaves the rest of the shape's
`<a:effectLst>` in place, written in the order `CT_EffectList` states — so a
shape can carry a shadow and a glow at once, as PowerPoint routinely writes
them. Setting the same effect twice still replaces it, because the schema
allows each one only once. `clearShapeEffects` remains the way to empty the
list.

`getShapePatternFill` now reports `preset` as the `PatternPreset` union rather
than a bare `string`, so what it reads can be handed straight back to
`setShapePatternFill`. An unrecognised `prst` in the file reads as `'pct50'`,
which is what PowerPoint paints when the attribute is absent.
