---
'@office-kit/pptx-editor': minor
'@office-kit/pptx-dev': patch
---

Agents in the browser can now read and edit the embedded editor's presentation through the `mountEditor` handle:

- `selection()` returns the shapes the user selected as `ShapeRef`s (`{ slideIndex, slide, shapeId, name }`), and `resolveShape(presentation, ref)` finds such a shape again, throwing once it has been deleted.
- `apply(label, edit)` runs `edit(presentation)` with the `@office-kit/pptx` API as one undo step named "Agent: label" (「エージェント: label」 in Japanese). An edit that throws is rolled back completely and `apply` rejects with its error.
- `on('change', listener)` reports every kept edit with its source (`'user'` or `'agent'`), including Undo and Redo; `on('selectionchange', listener)` reports the user's selection.
