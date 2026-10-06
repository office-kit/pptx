---
"@office-kit/pptx": minor
"@office-kit/pptx-dev": patch
---

Copy existing formatted text into a shape with `setShapeParagraphs(shape, { source, range })`. The optional UTF-16 range retains paragraph properties, fields, run formatting and link relationships while keeping the destination text body's settings.

Splitting an outline title now retains hyperlinks and paragraph formatting in the new title.
