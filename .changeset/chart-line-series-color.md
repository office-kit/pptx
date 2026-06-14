---
'pptx-kit': patch
---

fix: line-chart series colors now paint the line. The color was written only
as a bare `<a:solidFill>`, which doesn't color a line series' stroke, so
PowerPoint ignored it and fell back to its automatic palette (a 4-series line
chart authored as accent1–4 rendered blue/red/green/purple instead of the
requested colors). The color is now also emitted on `<a:ln>`.
