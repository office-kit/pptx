---
'pptx-kit': minor
---

feat: `getSlideBackgroundPatternFill(pres, slide)` returns the pattern
preset + theme-resolved foreground / background for slides whose
`<p:bgPr>` carries a `<a:pattFill>`. Playground now paints pattern
slide backgrounds via the same SVG `<pattern>` tile generator that
handles shape pattern fills.
