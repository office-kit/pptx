---
'pptx-kit': minor
---

feat: `getTableStyleFlags(table)` returns the `<a:tblPr>` boolean
toggles — `firstRow` / `lastRow` / `firstCol` / `lastCol` / `bandRow`
/ `bandCol`. Playground projects each onto a theme-derived tint
(accent1 for header / footer rows, 92%-white-mixed accent for bands)
when the cell doesn't supply an explicit fill of its own. Header text
rendered on the accent gets white text instead of the default body
color, matching PowerPoint's built-in table styles.
