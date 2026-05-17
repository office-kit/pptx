---
'pptx-kit': minor
---

feat: `getTableCellMargins(cell)` returns the cell's `<a:tcPr marL
marR marT marB>` inset margins in EMU. Each side is `null` when the
cell doesn't author it, so renderers know to fall back to
PowerPoint's defaults (91440 EMU / 0.1 in horizontal, 45720 EMU /
0.05 in vertical).
