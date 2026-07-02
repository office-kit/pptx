---
"pptx-kit": minor
---

Add `setShapeAdjustValues(shape, values)` to author a preset shape's adjust-handle guides (`<a:prstGeom><a:avLst>`). It's the mutating companion to the existing `getShapeAdjustValues` reader — pass raw ECMA-376 guide values keyed by guide name. The common use is the `roundRect` corner radius via the `adj` guide (`0..50000`; `setShapeAdjustValues(shape, { adj: 5000 })` gives a subtle 5% rounding). Throws when the shape has no preset geometry.
