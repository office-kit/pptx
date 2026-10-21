---
'@office-kit/pptx': patch
---

Charts now write their axis tick marks even when none are authored. An absent `<c:majorTickMark>` / `<c:minorTickMark>` means `cross` in the schema, so some apps drew ticks crossing every axis. The defaults written are: value axes (including a scatter / bubble chart's horizontal axis and the secondary value axis) major `out`, minor `none`; the category axis major `none` on bar and column charts and `out` on the others, minor `none`; the series (depth) axis of 3-D and surface charts major `out`, minor `none`. Authored `*TickMark` fields are written as before. Reading such a chart back now reports these tick marks.
