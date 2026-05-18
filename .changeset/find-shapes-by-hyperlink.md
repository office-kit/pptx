---
'pptx-kit': minor
---

feat: `findShapesByHyperlink(slide, url)` — slide-scoped finder that
returns every shape whose hyperlink target matches `url` (substring or
`RegExp`). Pairs the existing presentation-level
`findSlidesByHyperlink` for cases where the caller already has a
specific slide and wants the linking shapes inside it.
