---
'pptx-kit-preview': patch
---

Close three renderer correctness gaps surfaced by the expanded corpus:

- **Preset pattern fills** now render the real ECMA-376 `ST_PresetPatternVal`
  tokens (`horz`/`vert`/`ltHorz`/`ltVert`/`dotGrid` etc.) instead of falling
  through to a 50%-coverage checker — the old matcher keyed on GDI HatchStyle
  names no valid OOXML emits.
- **`wordArtVert` / `wordArtVertRtl`** stack glyphs upright (one per line) per
  `ST_TextVerticalType`, instead of rotating the run 90°, matching PowerPoint
  and the browser (`text-orientation:upright`) path.
- **`<a:normAutofit/>` without a baked `fontScale`** now shrinks text to fit the
  box, so overflowing bodies render at the reduced size PowerPoint/LibreOffice
  compute at display time rather than spilling past the box. The shrink factor is
  computed once and shared, so the server (SVG) and browser (`foreignObject`)
  previews agree.
