---
'pptx-kit': minor
---

feat: writers for the three stroke attributes that had readers but no
setters — `setShapeStrokeCap(shape, 'rnd' | 'sq' | 'flat' | null)`,
`setShapeStrokeJoin(shape, 'round' | 'bevel' | 'miter' | null)`, and
`setShapeStrokeCompound(shape, 'sng' | 'dbl' | 'thickThin' | 'thinThick' | 'tri' | null)`.

Cap and compound map to `<a:ln cap=…/>` and `<a:ln cmpd=…/>` attributes;
join writes one of the `<a:round/>` / `<a:bevel/>` / `<a:miter/>` child
variants. Passing `null` clears the attribute / removes the child so the
shape inherits the default. Creates `<a:ln>` if absent.
