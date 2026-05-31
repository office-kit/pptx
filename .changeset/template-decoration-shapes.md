---
'pptx-kit': minor
---

feat: add `getSlideLayoutShapes(pres, layout)` and `getSlideMasterShapes(pres,
layout)` — the non-placeholder decorative shapes (corner bars, divider lines,
logos, watermark text) on a slide layout and its master, as render-ready
`SlideShapeData`. Unlike the older flat `getSlideLayoutBackgroundShapes`, these
include pictures and groups and work with every `getShape*` reader, so a
picture logo's bytes resolve (against the layout/master's own relationship
table). For reading/rendering — the handles are bound to the layout/master
part, not a slide.
