---
'pptx-kit': minor
---

feat: `findShapeByName(slide, name)` now accepts a `RegExp` as well
as a literal string. Mirrors the RegExp support just landed on
`findShapesByName` (multi-match). Returns the first match in document
order; backward compatible with existing string callers.
