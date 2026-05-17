---
'pptx-kit': minor
---

feat: table span + border read-back.

- `getTableCellSpan(cell)` returns `{ gridSpan, rowSpan, hMerge, vMerge }`
  so renderers know which cells own a merged region and which are
  absorbed into one.
- `getTableCellBorders(pres, cell)` returns per-side borders (left,
  right, top, bottom, plus the two diagonals tlToBr / blToTr) with
  theme-resolved colors, widths, and dash style.

Playground table rendering now honours both: merged cells are skipped
on their absorbed positions, and per-cell borders render at the
authored color / width on top of the default thin grid.
