---
'@office-kit/pptx-dev': minor
---

`office-pptx dev` now merges TSX changes with canvas edits instead of asking you to choose between them. When the source is rewritten (for example by Claude Code) while you have saved or unsaved edits, changes to different shapes, slides, media and relationships are combined, and `build` / `exportDeck` export the merged deck. The editor only asks to **Keep my edits** or **Use source** when the same shape (or slide setting) was changed both in the editor and in the source, and it names what collided, e.g. "Slide 2: Title 1 was changed both here and in the source." Editor files saved by earlier versions are still read, but merge only after the next save.
