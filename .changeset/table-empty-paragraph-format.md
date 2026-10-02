---
'@office-kit/pptx': minor
'@office-kit/pptx-preview': patch
'@office-kit/pptx-dev': patch
---

Preserve empty table paragraphs' authored font size in the preview and text editor, so blank lines no longer collapse to a default size. The font controls now resolve inherited formatting at an empty cell paragraph's caret. Pass a null run index to getTableCellRunFormatEffective to read the effective paragraph end format.
