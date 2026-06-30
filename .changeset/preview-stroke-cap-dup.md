---
'pptx-kit-preview': patch
---

Fix a crash when rendering a line or connector that sets an explicit line
cap or join (`setShapeStrokeCap` / `setShapeStrokeJoin`). The renderer
emitted both its default `stroke-linecap="round"` and the shape's explicit
cap on the same element, producing a duplicate SVG attribute that aborted
the render ("attribute 'stroke-linecap' is already defined"). The default
is now only applied when the shape does not specify its own.
