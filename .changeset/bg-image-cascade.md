---
'pptx-kit': minor
---

feat: `getSlideLayoutBackgroundImageBytes(pres, layout)` and
`getSlideMasterBackgroundImageBytes(pres, layout)` complete the
picture-background cascade. The slide reader already returned bytes
for slide-level `<a:blipFill>` backgrounds; the new readers resolve
the same shape on layouts and masters via their own rel lists. The
playground renderer threads slide → layout → master fallback, so
template-defined photo backgrounds finally show on inheriting slides.
