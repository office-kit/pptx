---
'pptx-kit': minor
---

feat: more name-based finders now accept `RegExp` —
`findSlideLayout(pres, name)`,
`findCommentAuthorByName(pres, authorName)`, and
`findSlidesByLayoutName(pres, layoutName)`. Pairs the RegExp support
recently added to `findShapeByName` / `findShapesByName` /
`findCommentsByAuthor` / `findSlideByTitle`. String callers unchanged.
