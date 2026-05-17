---
'pptx-kit': minor
---

feat: `getTableStyleId(table)` returns the GUID string inside
`<a:tbl><a:tblPr><a:tableStyleId>`. PowerPoint references built-in
table styles (`{5C22544A-…}` = Medium Style 2 - Accent 1, etc.) and
theme-local styles by GUID. Returns `null` when the table doesn't
author one.
