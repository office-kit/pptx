---
'pptx-kit': minor
---

feat: `findChartsWithTrendlines(slide)` — slide-scoped finder for
charts that carry at least one `<c:trendline>` on any of their
series. Useful for deck-audit reports — trendlines are easy to add
and easy to forget. Charts whose kind isn't modeled are skipped.
