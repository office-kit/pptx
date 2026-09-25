---
"@office-kit/pptx": minor
"@office-kit/pptx-dev": patch
---

Choose pattern foreground and background colors from theme and standard palettes. Base theme colors stay linked when switching fill types and after saving.

`getShapePatternFill` accepts `preserveTheme: true` to return untransformed theme references; its default continues to return resolved RGB colors.
