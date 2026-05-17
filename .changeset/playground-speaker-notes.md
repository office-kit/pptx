---
'pptx-kit': minor
---

feat(site/playground): show speaker notes under each slide. The
playground now calls `getSlideNotes` for every slide and renders a
collapsible `<details>` block when notes exist, so users can
inspect the deck author's notes without opening PowerPoint.
