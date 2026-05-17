---
'pptx-kit': minor
---

feat: `ChartSpec.grouping` carries the `<c:grouping>` token —
`'clustered' | 'stacked' | 'percentStacked' | 'standard'`. Playground
column chart renders stacked / percent-stacked layouts: series stack
within each category, and percent-stacked normalises to 0..100% per
column with in-bar value labels.

Adds the `ChartGrouping` type to the public surface.
