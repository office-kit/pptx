---
"@office-kit/pptx": minor
"@office-kit/pptx-dev": patch
---

Add a bilingual table properties pane with cell selection, text, fill, alignment,
row and column dimensions, insertion and deletion. Save edits and restore them
through undo and reload.

Allow setTableCellText to preserve unaffected formatting. Read soft breaks and
field text in getTableCellText so editing does not silently drop visible content.
