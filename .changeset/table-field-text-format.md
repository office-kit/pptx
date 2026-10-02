---
"@office-kit/pptx": minor
"@office-kit/pptx-preview": patch
"@office-kit/pptx-dev": patch
---

Table fields such as dates and slide numbers now inherit cell text formatting in previews, text editing, and copied text. The existing getTableCellRunFormatEffective API accepts a fieldIndex selector to resolve field formatting.
