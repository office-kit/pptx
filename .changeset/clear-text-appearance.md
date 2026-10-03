---
'@office-kit/pptx': minor
'@office-kit/pptx-dev': minor
---

Add Clear text formatting to the preview's English and Japanese text and table-cell toolbars, with Ctrl/Command+Backslash for selected text. Clearing preserves hyperlinks and paragraph settings and supports Undo/Redo.

The existing `setShapeTextFormat` and `setTableCellTextFormat` APIs now accept `{ reset: true }` to restore inherited run appearance before applying a new format, optionally limited to a character range.
