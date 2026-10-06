---
'@office-kit/pptx': minor
'@office-kit/pptx-dev': patch
---

Add optional speaker-notes color resolution through the notes master theme and color map. The notes editor uses it for display while retaining literal color references for editing and round trips.

Expose notes line-break metadata and paragraph/break-aware editing so soft breaks survive save/load and undo.

Keep the notes caret after inserted paragraph breaks and preserve theme-based typing colors across consecutive empty paragraphs.
