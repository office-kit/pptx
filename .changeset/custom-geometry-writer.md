---
'@office-kit/pptx': minor
---

Add `setShapeCustomGeometry(shape, { paths })`, the writer counterpart of `getShapeCustomGeometry`. It replaces a shape's preset or custom geometry with `<a:custGeom>` paths (moveTo, lnTo, arcTo, quadBezTo, cubicBezTo, close), rounding coordinates to whole numbers and rejecting non-finite values, empty paths and paths that do not start with moveTo before touching the shape.
