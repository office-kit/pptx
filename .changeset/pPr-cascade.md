---
'pptx-kit': minor
---

feat: `getParagraphPropertiesEffective(pres, shape, p)` — paragraph-property
cascade mirroring the rPr one. Resolves alignment, left / right / first-line
indents, line spacing, paragraph spacing (before / after), and rtl through
the paragraph → text-body lstStyle → layout placeholder lstStyle →
master placeholder lstStyle → master txStyles chain.

The playground uses it as the primary source of paragraph properties so
placeholders inherit their default alignment / line-spacing / indent from
the layout / master, with any per-slide override winning on top.

Adds the `ParagraphProperties` type to the public surface.
