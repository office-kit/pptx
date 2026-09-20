---
"@office-kit/pptx-preview": patch
---

fix: transparent column and bar series appeared opaque in previews (#349)

Column and bar previews now respect series fill opacity in clustered, stacked, and percent-stacked charts. Transparent bases remain part of the stack, preserving the position of visible segments in waterfall-style charts.
