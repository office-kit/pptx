---
'pptx-kit': minor
---

feat: `getTableCellTextDirection(cell)` reads `<a:tcPr vert="…"/>` —
the same token set as `getShapeTextDirection` (`vert`, `vert270`,
`eaVert`, `mongolianVert`, `wordArtVert`, `wordArtVertRtl`).
Vertical column headers in tables commonly use `vert270` / `eaVert`
so the header label reads bottom-to-top alongside its column.
