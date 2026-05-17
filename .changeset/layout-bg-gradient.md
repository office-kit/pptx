---
'pptx-kit': minor
---

feat: `getSlideLayoutBackgroundGradientFill(layout)` returns the
gradient definition when a layout's background is
`<p:bgPr><a:gradFill>`. Same shape as the slide-level variant —
renderers can reuse the same projection logic for layout gradient
backgrounds via the shared `gradientDef` helper.
