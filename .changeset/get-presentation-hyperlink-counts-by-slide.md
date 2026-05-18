---
'pptx-kit': minor
---

feat: `getPresentationHyperlinkCountsBySlide(pres)` — dense per-slide
hyperlink count array. Counts shapes whose `getShapeHyperlink` is
non-null. Cheaper than `getAllHyperlinks` when the caller only wants
per-slide counts. Rounds out the deck-density family.
