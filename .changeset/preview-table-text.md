---
'pptx-kit-preview': patch
---

fix(preview): table cell text now renders with its real per-run formatting —
font size, bold, italic, color, typeface, and paragraph alignment — and wraps
within the cell width, in both the browser (`foreignObject`) and server
(`svg`) text modes. Previously every cell was drawn at a flat 18 pt with no
styling. Cells with no explicit run size still fall back to PowerPoint's 18 pt
default in the theme's body font.
