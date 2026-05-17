---
'pptx-kit': minor
---

feat: `getSlideLayoutBackground(layout)` mirrors `getSlideBackground`
for slide layouts. Playground falls back to it when the slide's own
background reports `'inherit'`, so brand-color or template backgrounds
authored on the layout actually paint behind slides that don't override
the bg themselves.
