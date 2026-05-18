---
'pptx-kit-site': patch
---

site(playground): show "comment authors" in the summary panel when
the deck has any review comments. Entries are sorted by count desc
and render as `Name (n) · …` — built on
`getPresentationCommentCountsByAuthor`.
