---
'pptx-kit': minor
---

feat: `getPresentationChartKindCounts(pres)` — deck-wide histogram of
`ChartKind` → count. Returns a frozen `Record` with every kind
present (zeros for absent kinds), so destructuring and chart-style
audits stay typed without runtime checks. Charts whose spec doesn't
parse are skipped, matching `findChartByKind` / `findSlidesWithChartKind`.
