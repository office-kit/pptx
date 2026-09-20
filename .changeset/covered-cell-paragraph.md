---
'@office-kit/pptx': patch
---

fix: Preserve valid table XML when setting alignment on a covered merged cell whose text was dropped. Recreated text bodies now include the required paragraph, so the alignment is applied and survives saving and reloading.
