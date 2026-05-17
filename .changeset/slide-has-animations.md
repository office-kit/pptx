---
'pptx-kit': minor
---

feat: `slideHasAnimations(slide)` — per-slide animation predicate.
Returns `true` when the slide carries a `<p:timing>` block (at least
one authored animation effect). Complements the deck-wide
`getPresentationSummary().hasAnimations`. The site playground uses
it (plus `getSlideTransition`) to show small `anim` / `trans`
badges next to each slide title so deck audits don't need to open
PowerPoint.
