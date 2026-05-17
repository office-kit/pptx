---
'pptx-kit': minor
---

feat: image color-effect readers — `isShapeImageGrayscale(shape)`
detects `<a:blip><a:grayscl/>` (Color > Grayscale), and
`getShapeImageBiLevelThreshold(shape)` returns the threshold percent
for `<a:blip><a:biLevel thresh="…"/>` (Color > Black and White).
Renderers can project these onto CSS / SVG filters.
