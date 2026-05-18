---
'pptx-kit': minor
---

feat: `findShapesByEffect(pres, slide, kind)` — returns every shape on
the slide whose `<a:effectLst>` carries an effect of the given `kind`
(`'outerShdw'`, `'innerShdw'`, `'glow'`, `'reflection'`, `'softEdge'`,
`'blur'`). Pure presence check; doesn't walk the layout / master
cascade. Useful for "which shapes have a shadow / glow on this
slide?" audits.
