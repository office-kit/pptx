---
'pptx-kit': minor
---

feat: add `isShapeTextBox(shape)` — `true` when a shape is a text box
(`<p:cNvSpPr txBox="1">`) rather than an autoshape. The two have different
default text formatting (text boxes left/top, autoshapes center/middle), so
renderers and layout code need to tell them apart.
