---
'pptx-kit': minor
---

feat: `findShapeInPresentation(pres, name)` now accepts a `RegExp` as
well as a literal string. Mirrors the RegExp support on the
slide-scoped `findShapeByName`. Backward compatible — string callers
still get exact-equality.
