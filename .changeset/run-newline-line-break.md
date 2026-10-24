---
'@office-kit/pptx': patch
---

Fixed a newline inside a run's text (`setShapeParagraphs`, `setTableCellParagraphs`, and the DSL `<Text paragraphs>`) being written as a raw newline inside `<a:t>`. LibreOffice read it as a paragraph break, so a bulleted item that wrapped onto two lines got a second bullet and the text overflowed its box. A `\n` or `\r\n` in a run is now written as a line break (`<a:br>`) between runs with the same format, which presentation apps and the preview all show as a line break inside the paragraph. Reading the text back still returns `\n` at each break; a `\r\n` reads back as `\n`.
