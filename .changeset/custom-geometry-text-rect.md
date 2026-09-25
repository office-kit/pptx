---
"@office-kit/pptx": minor
"@office-kit/pptx-preview": minor
"@office-kit/pptx-dev": patch
---

`getShapeCustomGeometry` now reports the `<a:rect>` a custom-geometry shape
states for its text, as `textRect`, with its guide formulas already evaluated.

The preview lays a custGeom shape's text in that rectangle instead of the whole
bounding box, through the new `shapeCustomTextRect` in `@office-kit/pptx-preview`,
and inline editing in the dev editor puts the caret in the same place. A custom
shape that states no rectangle still gets its whole box — a preset's
approximated region is never substituted for a shape that describes itself.
