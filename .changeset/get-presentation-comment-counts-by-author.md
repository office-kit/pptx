---
'pptx-kit': minor
---

feat: `getPresentationCommentCountsByAuthor(pres)` — deck-wide
histogram of comment counts keyed by author display name. Useful for
"who reviewed this deck the most?" audits. Authors sharing a display
name get merged into the same bucket; pair with
`getPresentationCommenters` when authors with identical names need to
be kept separate by `id`.
