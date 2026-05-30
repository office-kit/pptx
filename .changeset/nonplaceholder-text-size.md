---
'pptx-kit': patch
---

fix: `getShapeRunFormatEffective` / `getParagraphPropertiesEffective` no longer
inherit the slide master's `bodyStyle` for plain text boxes. A shape without a
`<p:ph>` is not a placeholder, so its unsized runs now resolve to no inherited
size (consumers apply the ~18pt text-box default) instead of wrongly picking up
the master body size (often much larger). Real placeholders — including ones
whose `<p:ph>` omits a `type` — still inherit as before. This makes effective
text formatting match what PowerPoint and LibreOffice render for text boxes.
