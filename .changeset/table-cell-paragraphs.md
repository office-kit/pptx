---
'pptx-kit': minor
---

feat: `getTableCellParagraphs(cell)` returns a table cell's text as structured
paragraphs — each carrying its alignment and per-run format (`size`, `bold`,
`italic`, `color`, `font`, …) — the rich counterpart to `getTableCellText`,
which only returns the flat visible string.
