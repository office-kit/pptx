---
'pptx-kit': minor
---

feat: `findShapesInRect(slide, x, y, w, h)` — marquee-style region
finder. Returns every shape whose bounds overlap the rectangle
(touching edges count). Shapes with no resolvable bounds are skipped.
Companion to `findShapesAtPoint(slide, x, y)` for cases where the
caller wants a region of the slide rather than a single point.
