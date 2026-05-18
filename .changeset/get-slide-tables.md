---
'pptx-kit': minor
---

feat: `getSlideTables(slide)` — returns every table graphic-frame
shape on the slide, in document order. Pair to `getSlideCharts` for
cases where the caller wants just the tables; convenience over
`getSlideShapes(slide).filter(isTableShape)`.
