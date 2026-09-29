---
'@office-kit/pptx-preview': minor
---

feat: the SVG ties clicks back to the deck — each slide shape (group members included) is wrapped in a `<g data-pptx-shape-id>` carrying `getShapeId`, paragraphs carry `data-pptx-paragraph`, and table cells carry `data-pptx-cell="row,col"`. Rendering is otherwise unchanged.
