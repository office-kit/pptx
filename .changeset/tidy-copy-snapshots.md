---
"@office-kit/pptx": minor
"@office-kit/pptx-dev": patch
---

Allow `copyShape` to copy between presentations while preserving related images, charts, embedded workbooks and unknown dependencies. Allocate unique IDs for every shape inside copied groups and preserve their internal connector references.

Fix cut and paste in the development preview. Copied content now retains its state even after editing or deleting the source or opening another document, and undo/redo restores the pasted selection.
