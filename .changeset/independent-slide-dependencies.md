---
"@office-kit/pptx": patch
---

Duplicated slides now have independent charts, embedded workbooks and notes.
Editing a duplicate no longer changes the original's chart or notes. Unknown
owned dependency parts are copied too, including cyclic relationships. Layouts,
masters, themes and media remain shared. Invalid missing dependencies fail before
parts are added.
