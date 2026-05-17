---
'pptx-kit': minor
---

feat: `getShapeParagraphElements(shape, paragraphIndex)` returns the
inline children of a paragraph (runs, field placeholders, and line
breaks) in document order. Renderers can walk this list to reproduce
the full visible content — footer / date / slide-number `<a:fld>`
text was previously dropped by the strict `<a:r>`-only run accessors.

The playground now uses it: footer text + slide numbers + datetime
fields show up in the preview, and `<a:br>` line breaks render as
real `<br/>` inside the foreignObject body.

Adds the `ShapeParagraphElement` discriminated union to the public
type surface.
