---
"@office-kit/pptx": minor
"@office-kit/pptx-dev": minor
---

Add a bilingual find-and-replace dialog to the development preview, with match navigation, individual and bulk replacement, case sensitivity, current-slide scope, table cell support, keyboard shortcuts and undoable changes. Replacement values are literal text.

Add an optional UTF-16 `range` to `setShapeText` and `setTableCellText` to replace an exact selection while retaining unaffected formatting, including when adjacent characters are identical. Invalid boundaries and split surrogate pairs are rejected.
