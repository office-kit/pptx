---
'pptx-kit': minor
---

feat(site/playground): paragraphs with image bullets (`<a:pPr><a:buBlip>`)
render a filled-square glyph (■) instead of inheriting the default
round bullet. The reader already exposed `isParagraphBulletPicture`;
the playground now threads it through paragraph metadata so the
visual cue lands.
