---
'pptx-kit': minor
---

feat: `ChartSpec.categoryAxisNumberFormat` — number-format code for the
category-axis tick labels (`<c:catAx><c:numFmt formatCode="…"/>`). Most
useful on date-style categories (`"mm/dd/yyyy"`, `"mmm-yyyy"`) but
accepts any Excel format string. Independent of `valueAxis.numberFormat`.
Read by chart-reader, written by chart-builder in the correct CT_CatAx
schema order (after `<c:title>`, before `<c:majorTickMark>`).
