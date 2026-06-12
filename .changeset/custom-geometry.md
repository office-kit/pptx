---
'pptx-kit': minor
'pptx-kit-preview': minor
---

feat: read custom geometry. New `getShapeCustomGeometry(shape)` returns a
shape's `<a:custGeom>` (ECMA-376 §20.1.9) as a fully-evaluated path list —
guide formulas (`avLst`/`gdLst`, all §20.1.9.11 operators) are resolved
against the shape extents so the returned `moveTo`/`lnTo`/`arcTo`/
`quadBezTo`/`cubicBezTo`/`close` commands carry only numbers. The preview
renderer now draws custom geometry as a real SVG path (arcs converted to
cubic Béziers) instead of a labelled rectangle placeholder; only a custGeom
that fails to evaluate still falls back, marked `data-pptx-fallback`.
