---
'pptx-kit': minor
---

feat(chart): axis number formats now accept Excel's `"$"#,##0`
quoted-literal prefix / suffix syntax. PowerPoint typically emits
currency as `"$"#,##0` (or `"\$"#,##0`) rather than the bare `$`
form, so the previous detection missed it.
