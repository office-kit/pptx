---
'pptx-kit': minor
---

feat: `findShapesByName(slide, name)` now accepts a `RegExp` as well
as a literal string. Useful when template-cloned shapes share a
prefix (`'TextPlaceholder1'`, `'TextPlaceholder2'`, …). Backward
compatible — string callers still get exact-equality matching.
