---
'pptx-kit': minor
---

feat: `setShapeTextDirection(shape, direction | null)` — companion
writer for the existing `getShapeTextDirection` reader. Sets
`<a:bodyPr vert="…"/>` with any of the six `ST_TextVerticalType`
values (`vert`, `vert270`, `wordArtVert`, `eaVert`, `mongolianVert`,
`wordArtVertRtl`); passing `null` or `'horz'` clears the attribute so
the shape uses the default horizontal direction.
