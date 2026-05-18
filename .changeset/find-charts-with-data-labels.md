---
'pptx-kit': minor
---

feat: `findChartsWithDataLabels(slide)` — slide-scoped auditor for
charts whose chart-level or per-series `dataLabels` enable at least
one of `showValue` / `showCategory` / `showSeriesName` / `showPercent`.
Purely presence-based; doesn't validate numberFormat or position.
Charts whose kind isn't modeled are skipped.
