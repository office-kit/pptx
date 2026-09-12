---
'@office-kit/pptx-preview': patch
---

Fix missing EMF artwork and abrupt background color changes in imported slide previews.

EMF pictures made of solid-filled line and Bézier paths now render as transparent
vector images, including artwork inherited from slide layouts. Images with
unsupported drawing commands retain their placeholder instead of rendering only
part of the artwork. Gradient colors are ordered by their positions before
rendering, so out-of-order stops no longer introduce flat bands or color jumps.
