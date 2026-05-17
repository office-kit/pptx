---
'pptx-kit': minor
---

feat(site/playground): apply the 3-level gradient bg cascade. When
the slide reports `'gradient'` but doesn't author the actual stops,
the renderer walks slide → layout → master gradient-fill readers
to find the inherited gradient definition.
