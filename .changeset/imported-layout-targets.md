---
'@office-kit/pptx': patch
---

Preserve the actual package path when switching slide layouts, including imported layouts outside the conventional slideLayouts folder. This prevents broken or incorrect layout references after switching and saving.
