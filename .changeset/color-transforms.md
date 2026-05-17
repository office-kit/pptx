---
'pptx-kit': minor
---

feat: apply ECMA-376 §20.1.2.3.x color transforms when resolving colors.

- New `resolveDrawingColor(colorEl, theme)` resolves any DrawingML color
  element (`<a:srgbClr>` / `<a:schemeClr>` / `<a:sysClr>` / `<a:prstClr>`)
  with all transform children (`<a:lumMod>`, `<a:lumOff>`, `<a:shade>`,
  `<a:tint>`, `<a:satMod>` / `Off`, `<a:hueMod>` / `Off`, `<a:gray>`,
  `<a:inv>`, `<a:comp>`) applied. Scheme tokens are looked up against
  the supplied theme.
- New `getShapeFillColorResolved(pres, shape)` and
  `getShapeStrokeColorResolved(pres, shape)` return the exact `#RRGGBB`
  PowerPoint paints — useful for renderers / exporters where the legacy
  `getShapeFillColor` / `getShapeStrokeColor` strings (`#RRGGBB` or
  `scheme:<token>`) miss both scheme resolution and color transforms.
- `getShapeRunFormatEffective` now applies the same pipeline at every
  layer of the rPr cascade, so a run inheriting `accent1 lumMod=40000
  lumOff=60000` (PowerPoint's "Accent 1, Lighter 60%") resolves to the
  concrete tinted hex instead of leaking the raw token through.
