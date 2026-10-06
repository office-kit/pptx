---
'@office-kit/pptx': patch
---

`setShapeText`, `setTableCellText` and `setSlideNotes` with a `range` no longer leave a separate `<a:r>` for every insertion: text inserted next to a run with the same formatting extends that run, and deleting the text between two halves of a run joins them again, as PowerPoint does. Typing in the editor previously saved one run per keystroke.
