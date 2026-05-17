---
'pptx-kit': minor
---

feat: Tier B fidelity batch.

- `getShapeTextDirection(shape)` returns the `<a:bodyPr vert="…"/>`
  token (`vert`, `vert270`, `wordArtVert`, `eaVert`, `mongolianVert`,
  `wordArtVertRtl`). Playground projects each onto a CSS
  `writing-mode` / `text-orientation` declaration so Asian and
  Mongolian-style vertical text renders without manual transforms.
- Playground wraps shapes carrying a `<a:hlinkClick>` in an SVG `<a>`
  element so the preview is clickable — matches PowerPoint's
  slide-show behaviour for shape-level hyperlinks.
- Group shape rendering now applies the group's own `<a:xfrm rot
  flipH flipV>` to the whole subtree before the scale + translate
  that maps internal coords onto slide coords.
