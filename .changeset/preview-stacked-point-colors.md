---
"@office-kit/pptx-preview": patch
---

fix: highlighted points lost their colors in stacked chart previews (#350)

Stacked and percent-stacked column and bar previews now preserve per-point highlight colors. They use the same point-color, single-series varying-color, series-color, and palette fallbacks as clustered charts.
