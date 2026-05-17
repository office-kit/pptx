---
'pptx-kit': minor
---

feat: `getSlideLayoutBackground` now handles `<p:bgRef>` the same way
`getSlideBackground` does. Layouts in real brand templates almost
always reference the theme via `<p:bgRef>` rather than authoring an
explicit `<p:bgPr>` — picking up the inner color element as a solid
fill closes the cascade so the playground paints the right brand
color even when the slide's own background reports `inherit`.
