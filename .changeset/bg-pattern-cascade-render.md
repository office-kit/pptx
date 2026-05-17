---
'pptx-kit': minor
---

feat(site/playground): apply the 3-level pattern bg cascade. When the
slide reports `'pattern'` but doesn't author the actual pattern preset,
the renderer now walks slide → layout → master pattern-fill readers,
paralleling the gradient cascade.
