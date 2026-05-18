---
'pptx-kit': minor
---

feat: `getPresentationCommentCountsBySlide(pres)` — dense per-slide
comment count array. Every slide appears as an element (count `0`
when the slide has no comments), so callers can chart comment
density per slide without re-indexing.
