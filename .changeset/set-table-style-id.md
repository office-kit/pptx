---
'pptx-kit': minor
---

feat: `setTableStyleId(table, styleId | null)` — writer for
`<a:tbl><a:tblPr><a:tableStyleId>`. Pairs the existing `getTableStyleId`
reader. Pass the curly-braced GUID (e.g.
`'{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}'` for PowerPoint's "Medium
Style 2 - Accent 1") or `null` to remove the reference so the table
uses the slide's default style. Creates `<a:tblPr>` if absent. Throws
when the shape isn't a table graphic frame.
