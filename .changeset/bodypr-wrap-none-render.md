---
'pptx-kit': minor
---

feat(site/playground): shape text honors `<a:bodyPr wrap="none"/>`.
The reader's effective wrap value was already threaded through; the
renderer now emits `white-space:nowrap` when wrap is `'none'`,
keeping single-line text frames (vertical labels, breadcrumbs,
fixed-width badges) from wrapping into multiple lines.
