---
'pptx-kit': minor
---

feat: chart builder writes back chart-level data-label config. A new
`dLblsElement` helper builds `<c:dLbls>` with `showVal` / `showCatName`
/ `showSerName` / `showPercent` toggles plus optional `<c:numFmt>`,
`<c:dLblPos>`, and `<c:separator>`. Wired into every chart variant
(bar / column / line / pie / doughnut / area), so round-tripping a
chart with authored data labels preserves them.
