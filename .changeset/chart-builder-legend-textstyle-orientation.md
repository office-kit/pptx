---
'pptx-kit': minor
---

feat: chart builder writes back legend.textStyle + axis orientation
reversals. `<c:legend><c:txPr>` now carries the authored font / color
from `legend.textStyle` (via the existing `axisTxPrElement` helper);
`<c:scaling><c:orientation>` honors `categoryAxisOrientation` and
`valueAxisOrientation` (defaulting to `minMax`). Round-trip test
asserts all four survive read → save → reload.
