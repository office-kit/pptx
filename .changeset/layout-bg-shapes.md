---
'pptx-kit': minor
---

feat: `getSlideLayoutBackgroundShapes(pres, layout)` returns the
non-placeholder shapes on a layout as a render-ready view
(`SlideLayoutBackgroundShape` — bounds, preset, fillHex, strokeHex,
strokeWidthEmu, rotation, flip). Playground paints them behind the
slide's own shapes so brand-template decoration (corner bars, divider
lines, background rectangles) shows through on slides that don't
redefine the layout themselves.

Adds the `SlideLayoutBackgroundShape` type to the public surface.
