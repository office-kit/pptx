---
'pptx-kit': minor
---

feat: `ChartSpec.categoryAxisNoMultiLevelLabel` — toggle multi-level
(hierarchical) category labels via `<c:catAx><c:noMultiLvlLbl val/>`.
PowerPoint defaults to `0` (multi-level labels stack); set to `true`
to flatten hierarchical categories into a single row. Read by
chart-reader, written by chart-builder at the schema-required last
position inside `<c:catAx>`.
