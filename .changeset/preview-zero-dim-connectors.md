---
'pptx-kit-site': patch
---

site(preview): render horizontal and vertical connectors. Lines have a
zero-height (horizontal) or zero-width (vertical) bounding box, which the
shape renderer's degenerate-shape guard was discarding — so straight lines and
arrows simply didn't appear. Connectors now render unless they are a true point.
