---
'pptx-kit': minor
---

feat(site/playground): render `<a:duotone>` image recolor. The filter
pipeline desaturates the picture to luminance, then samples a
two-color gradient (firstColor → secondColor) via a 16-step
`feComponentTransfer` table. Pictures with PowerPoint's Color >
Recolor preset finally render in their authored two-color tint.
