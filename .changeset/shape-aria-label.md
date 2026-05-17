---
'pptx-kit': minor
---

feat(site/playground): shape `aria-label` from authored alt text.
Each rendered shape with a non-empty alt title (or, as fallback,
alt description) now exposes `role="img" aria-label="…"` on the
root `<g>`. Screen readers announce decks the same way PowerPoint's
Accessibility Inspector reports them, without affecting visuals.
