---
'pptx-kit': minor
---

feat: `getParagraphIndent(shape, p)` returns the paragraph's
`<a:pPr marL marR indent>` values in EMU (`null` for sides the
paragraph doesn't author). Playground projects each side to CSS
`padding-left` / `padding-right` / `text-indent` and skips the
level-based default when the paragraph carries an explicit `marL`.
