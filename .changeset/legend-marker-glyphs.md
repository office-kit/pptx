---
'pptx-kit': minor
---

feat(site/playground): line / area chart legend swatches use the
series marker glyph. The legend previously rendered every series as
a 9×9 square color rect. For `line` / `area` charts the renderer now
passes the per-series `markerSymbol` (circle / square / diamond /
triangle / star / x / plus / dash / dot) so legend entries match
the data points. Bar / column / pie keep the square swatch.
