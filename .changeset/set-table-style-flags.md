---
'pptx-kit': minor
---

feat: `setTableStyleFlags(table, flags)` — partial-update writer for
the six `<a:tblPr>` boolean style flags (`firstRow`, `lastRow`,
`firstCol`, `lastCol`, `bandRow`, `bandCol`). Pairs the existing
`getTableStyleFlags` reader. Only the keys present in `flags` are
touched — omitted keys keep their current state. A flag set to `false`
strips the attribute (matching how PowerPoint round-trips defaults).
Creates `<a:tblPr>` if absent. Throws when the shape isn't a table
graphic frame.
