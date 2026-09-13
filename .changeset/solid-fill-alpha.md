---
'@office-kit/pptx': minor
'@office-kit/pptx-preview': patch
---

Read and render the transparency of solid fills and outlines.

`@office-kit/pptx` adds `getShapeFillOpacity`, `getShapeStrokeOpacity`, and
`resolveDrawingColorOpacity`, which resolve `<a:alpha>` / `<a:alphaMod>` /
`<a:alphaOff>` on a color to a 0–1 opacity. `getShapeFillColorResolved` and
`getShapeStrokeColorResolved` keep returning the plain `#RRGGBB`.

`@office-kit/pptx-preview` now emits `fill-opacity` / `stroke-opacity` for
translucent solid fills and outlines, including artwork inherited from slide
layouts and masters. A semi-transparent shape layered over a gradient or the
slide background previously rendered as an opaque block that hid whatever sat
beneath it.
