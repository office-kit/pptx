---
'pptx-kit': minor
---

feat: `getShapeGradientFill` now surfaces non-linear gradient paths
(`<a:path path="circle|rect|shape">`) and the `<a:fillToRect>` focus
rectangle. `GradientFillOptions` gains `path` and `focus` fields so
renderers can reproduce radial, rectangular, and shape-following
gradients instead of falling back to a linear approximation.

The playground renderer emits an SVG `<radialGradient>` for the
non-linear paths, with reversed stop offsets so the first ECMA-376
stop sits at the focus center (matching PowerPoint's outward
painting order).
