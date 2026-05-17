---
'pptx-kit': minor
---

feat: `getTableCellAnchor(cell)` returns the cell's vertical text
anchor (`<a:tcPr anchor="t|ctr|b"/>`) as `'top' | 'center' |
'bottom' | null`. Playground projects each onto a CSS
`justify-content` so cell text sits at the authored vertical
position instead of always centering.
