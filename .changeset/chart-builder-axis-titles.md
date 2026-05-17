---
'pptx-kit': minor
---

feat: chart builder writes back axis titles, hidden flags, and
category-axis tick-label config. `<c:valAx>` / `<c:catAx>` now emit:

- `<c:title>` with style (from `valueAxisTitleStyle` /
  `categoryAxisTitleStyle`) when an axis title is authored
- `<c:delete val="1"/>` when `valueAxisHidden` / `categoryAxisHidden`
- `<c:tickLblPos>` and `<c:tickLblSkip>` when authored on the
  category axis

Closes the read/write gap for these `ChartSpec` fields. Round-trip
test added.
