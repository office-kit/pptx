---
"@office-kit/pptx": minor
"@office-kit/pptx-preview": minor
"@office-kit/pptx-dev": minor
---

Text can now carry its own outline, shadow and glow — the WordArt half of a
character format.

`TextFormat` gained `outline`, `shadow` and `glow`, so every writer that takes
one (`setShapeTextFormat`, `setShapeRunFormat`, `setTableCellTextFormat`, …)
writes `<a:ln>` and `<a:effectLst>` into the run's `<a:rPr>`, and
`getShapeRunFormat` / `getShapeRunFormatEffective` read them back. Passing
`null` removes one effect and leaves the others alone. `GlowOptions` also
gained the `opacity` its reader already reported, so a glow with `<a:alpha>`
round-trips instead of losing it.

The preview paints all three, and the editor exposes them in the text-format
dialog and the properties panel in English and Japanese, including through the
format painter. In the rasterised SVG text path only the outline is painted so
far.

`<Text>` in `@office-kit/pptx-dsl` keeps `shadow` and `glow` as the text box's
own effects; per-run effects go through `paragraphs`.
