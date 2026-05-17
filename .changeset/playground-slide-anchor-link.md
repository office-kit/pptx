---
'pptx-kit': minor
---

feat(site/playground): make the per-slide number an anchor link.
Each slide's two-digit index in the head row is now an `<a
href="#slide-N">` link, so users can right-click → "Copy link
address" to share a deep link to a specific slide. Paired with the
`id="slide-N"` already on each `<li>`, the link also scrolls the
slide into view when clicked.
