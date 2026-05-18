---
'pptx-kit': minor
---

feat: `setTableCellTextDirection(cell, direction | null)` — vertical-
text writer for table cells, paired with the existing
`getTableCellTextDirection` reader. Same six `ST_TextVerticalType`
values as `setShapeTextDirection`. Passing `null` (or `'horz'`) clears
the `<a:tcPr vert="…"/>` attribute so the cell uses the default
horizontal direction. Creates `<a:tcPr>` if absent.
