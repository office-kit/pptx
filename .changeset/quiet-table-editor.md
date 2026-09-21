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

Apply fill, alignment, bold and italic to selected cell ranges in a single undo
step, including merged cells, from the English and Japanese preview.

Use the shared text toolbar for table ranges, including font families, size,
color and underline. Display saved color and mixed sizes when selection changes.

Edit table border colors, widths and styles, with all-cell and outside-border
placement. Reset borders and undo the changes from the bilingual preview.

Double-click table cells directly on the canvas to edit their text and format
selected text, with merged-cell hit areas, visible cell selection, save and undo.
Allow setTableCellTextFormat to accept an optional UTF-16 range.
