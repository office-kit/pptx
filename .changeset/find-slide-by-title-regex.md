---
'pptx-kit': minor
---

feat: `findSlideByTitle(pres, title)` now accepts a `RegExp` as well
as a literal string. Pairs the RegExp support on
`findSlidesByText` / `findShapeByName` / `findCommentsByAuthor`.
Backward compatible — string callers still get exact-equality.
