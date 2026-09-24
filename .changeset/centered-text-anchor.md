---
'@office-kit/pptx': minor
'@office-kit/pptx-preview': minor
'@office-kit/pptx-dev': patch
---

Support PowerPoint's Top, Middle and Bottom Centered text anchors through the existing `setShapeTextAnchor` API's `centered` option. Resolve inherited centering through `getShapeBodyPrEffective`, preserve paragraph alignment, and expose all six anchor choices in the editor with preview and text-editing support.

Preview integrations can use `shapeTextAnchorOffset` to position editable text consistently with the rendered text block.
