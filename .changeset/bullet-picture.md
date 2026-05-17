---
'pptx-kit': minor
---

feat: `isParagraphBulletPicture(shape, p)` returns `true` when the
paragraph uses an image as its bullet (`<a:pPr><a:buBlip>`).
Renderers without image-bullet support can fall back to a generic
glyph; UIs that want to indicate the bullet is custom have a clean
yes/no signal.
