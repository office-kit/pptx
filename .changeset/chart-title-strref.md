---
'pptx-kit': minor
---

feat: chart titles read `<c:tx><c:strRef>` workbook-cell references.
Previously only literal `<c:rich>` titles surfaced; titles authored
via Excel's "Link to source cell" wizard (which emits `<c:strRef>`
with a `<c:strCache>` of the resolved text) now flow through to
`ChartSpec.title` as the cached value. Affects the title shown above
the chart and, transitively, axis-title rendering.
