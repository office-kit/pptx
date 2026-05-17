---
'pptx-kit': minor
---

feat: playground now applies the picture corrections that already
shipped on the API: source-rectangle crop (`<a:srcRect>`), brightness
(`<a:lumOff>`), contrast (`<a:lumMod>`), and opacity (`<a:alphaModFix>`).

Crops project to an enlarged `<image>` element clipped to the shape's
bounds (matching PowerPoint's "Crop" tool). Brightness + contrast
compose into an SVG `<feComponentTransfer>` filter. Opacity drives
the `opacity` attribute directly.
