---
'pptx-kit': minor
---

feat(chart): `ChartAxisScaling.numberFormat` reads `<c:valAx><c:numFmt
formatCode="…"/>`. Playground projects the most common Excel format
codes to axis labels — percent (`'0%'`, `'0.0%'`), thousand
separator (`'#,##0'`, `'#,##0.0'`), and currency prefixes
(`'$#,##0'`, `'¥#,##0'`). Other codes fall through to the generic
auto-formatted label.
