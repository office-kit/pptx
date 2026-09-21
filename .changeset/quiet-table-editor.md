---
"@office-kit/pptx": minor
"@office-kit/pptx-dev": patch
---

Add a bilingual table properties pane with cell selection, text, fill, alignment,
row and column dimensions, insertion and deletion. Save edits and restore them
through undo and reload.

Allow setTableCellText to preserve unaffected formatting. Read soft breaks and
field text in getTableCellText so editing does not silently drop visible content.

Merge selected table cells in the preview while retaining their formatted text,
and split merged cells. Expose splitTableCell and an append text policy for
mergeTableCells; preserve text and reject malformed merges before mutation.
