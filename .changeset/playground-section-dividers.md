---
'pptx-kit': minor
---

feat(site/playground): render section dividers in the slide list.
Reads `getSlideSections(pres)`, maps each section's first slide to
the section's name, and renders a dashed divider above that slide
in the slide list. Deck audits can now see the section grouping at
a glance.
