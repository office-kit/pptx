---
'pptx-kit': minor
---

feat: `findCommentsByAuthor(pres, authorName)` and
`findSlidesWithCommentsByAuthor(pres, authorName)` now accept a
`RegExp` as well as a literal string. Useful for "every comment from
review bots" (`/^review-bot/`) or "every comment from anyone with a
given email domain" patterns. Backward compatible — string callers
still get exact-equality matching.
