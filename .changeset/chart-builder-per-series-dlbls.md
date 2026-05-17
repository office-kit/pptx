---
'pptx-kit': minor
---

feat: chart builder writes back per-series `<c:dLbls>` overrides.
`dLblsElement` is refactored to take the labels arg directly via
`buildDLblsFromLabels(dl)`; `seriesElement` now emits per-series
`<c:dLbls>` when authored, so charts with per-series label toggles /
numberFormat / position survive round-trip. Round-trip test covers
all four fields plus the no-override case.
