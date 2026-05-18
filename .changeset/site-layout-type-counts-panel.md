---
'pptx-kit-site': patch
---

site(playground): show "layout types in use" in the summary panel
when slides reference at least one typed layout. Built on the new
`getSlideLayoutUsageCountsByType` aggregator; renders e.g.
`3 title · 12 obj · 1 blank`.
