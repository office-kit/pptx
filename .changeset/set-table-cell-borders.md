---
'pptx-kit': minor
---

feat: `setTableCellBorders(cell, sides | null)` — partial-update writer
for all 6 cell-border slots (`left`, `right`, `top`, `bottom` + the
`tlToBr` / `blToTr` diagonals). Pairs the existing
`getTableCellBorders` reader. Sides listed with `null` are removed from
`<a:tcPr>`; sides omitted are left untouched. Passing `null` as the
whole `sides` arg clears every side. Creates `<a:tcPr>` if absent.

The diagonals are independent of the four cardinal sides — a
strikethrough cell can have only `tlToBr`.
