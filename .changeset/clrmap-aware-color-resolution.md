---
'pptx-kit': minor
'pptx-kit-preview': minor
---

Resolve scheme colors through the slide's color map so inverted-map templates render correctly

Templates whose slide master inverts the color map (`<p:clrMap bg1="dk1" tx1="lt1">`, common in Google Slides / Canva exports) previously rendered with swapped light/dark colors: slide backgrounds came out black in the preview while PowerPoint paints them white, and generated tables and charts came out with invisible text (the default `tx1` token resolved to the same color as the background).

- **`getEffectiveColorMap(slide)`** — new export returning the slide's effective color map (the master's `<p:clrMap>` overlaid by a per-slide `<p:clrMapOvr>`). Color resolution and renderers apply it to `schemeClr` tokens before indexing the theme.
- **`resolveDrawingColor(colorEl, theme, clrMap?)`** — accepts an optional color map; scheme tokens are remapped through it before the theme lookup. Omitting it preserves the previous behavior (correct for the standard map).
- **`addSlideTable` / `addSlideChart`** now bake the deck's resolved body-text color onto table cells and chart text (axis labels, legend, data labels) so generated tables and charts stay readable regardless of the template's color map. Authored colors still win; override table cells afterwards with `setTableCellTextFormat`.
- **`pptx-kit-preview`** resolves `schemeClr` tokens through the effective color map, so previews of inverted-map decks match what PowerPoint paints.
