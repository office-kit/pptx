---
'@office-kit/pptx': patch
'@office-kit/pptx-preview': patch
---

Resolve text fonts from each slide's own master theme, and honor shape-style font and color defaults ahead of inherited placeholder formatting. Shapes using their group's fill now inherit solid colors, transparency, and gradient details instead of falling back to their own theme style.
