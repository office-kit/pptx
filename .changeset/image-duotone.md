---
'pptx-kit': minor
---

feat: `getShapeImageDuotone(pres, shape)` reads the picture's
`<a:blip><a:duotone>` two-color recolor effect — the typical
"Picture Tools > Recolor" output. Returns the two hex-resolved
colors (or `null` for each that the duotone didn't author). Lets
downstream renderers project the duotone via SVG `<filter>` or
inform consumers that the picture has a color-replacement applied.
