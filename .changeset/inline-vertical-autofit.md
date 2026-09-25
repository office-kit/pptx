---
"@office-kit/pptx-preview": minor
"@office-kit/pptx-dev": patch
---

Inline editing now reads vertical text the way the preview paints it, and
shrinks autofit text by the same factor — a `<a:normAutofit/>` title no longer
jumps back to its authored size the moment the caret appears.

Two new exports carry the shared rules: `verticalTextStyle` / `textColumnsStyle`
turn `<a:bodyPr vert=… numCol=… spcCol=…>` into the CSS both surfaces use, and
`shapeAutoFitScale` reports the shrink factor the renderer applies to a shape,
for the box it is actually laid out in.
