---
'pptx-kit': minor
---

feat: density-array companions for tables and images —
`getPresentationTableCountsBySlide(pres)` and
`getPresentationImageCountsBySlide(pres)`. Both return a dense
per-slide count array (0 for slides without that asset kind),
matching the shape / chart / comment / text-length counterparts.
Completes the deck-density family.
