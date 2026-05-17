---
'pptx-kit': minor
---

feat: chart titles honor `<a:rPr>` font / color overrides.
`ChartSpec.titleStyle` carries the authored size (in pt), bold, italic,
and fill color extracted from the title's first `<a:r><a:rPr>` (or
`<a:pPr><a:defRPr>` as fallback). The playground renderer projects
those through to the SVG `<text>`. Templates that brand their chart
titles to a non-default size / color finally render with the authored
look.
