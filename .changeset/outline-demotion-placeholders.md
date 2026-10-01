---
"@office-kit/pptx": minor
"@office-kit/pptx-dev": patch
---

Allow `addSlidePlaceholder` to inherit a placeholder directly from the slide master without changing the slide layout.

The editor now confirms deletion of additional objects when demoting an outline title, and restores a missing body placeholder even on Title Only slides. Cancel keeps the document unchanged; Undo restores the deleted slide and its objects.
