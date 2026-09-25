---
"@office-kit/pptx": minor
"@office-kit/pptx-dev": patch
---

Allow setShapeTextFormat to target a UTF-16 text range while retaining surrounding
run and paragraph properties. Reject invalid ranges without changing the document.

Add a bilingual selected-text formatting bar to the development preview with
bold, italic, underline, fonts, size and color, backed by persisted edits and undo.

Keep property controls synchronized with canvas edits and preserve mixed text
formatting when editing in the properties pane. Build the preview editor into
the development package regardless of the working directory.
