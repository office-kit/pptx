---
'pptx-kit': minor
---

feat: chart value-axis exposes `<c:scaling><c:logBase val="N"/>`.
`ChartAxisScaling.logBase` carries the authored log base (commonly
`2`, `10`, or `Math.E`). The reader clamps to PowerPoint's `[2, 1000]`
range. Callers that round-trip charts now preserve the log-scale
flag; the playground renderer still draws linear (log-scale
projection is a follow-up — exposing the field unblocks it).
