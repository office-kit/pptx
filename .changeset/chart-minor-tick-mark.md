---
'pptx-kit': minor
---

feat: `ChartSpec.valueAxisMinorTickMark` and `categoryAxisMinorTickMark`
— minor-tick-mark mode siblings of the existing `*MajorTickMark` pair.
Maps to `<c:catAx><c:minorTickMark val="in|out|cross|none"/>` and the
value-axis equivalent. Read by chart-reader, written by chart-builder
in the correct schema order (right after `<c:majorTickMark>`).
