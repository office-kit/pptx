---
'@office-kit/pptx-preview': patch
---

fix: bar and column charts ignored per-point colors (`pointColors` / `<c:dPt>`), so a "one bar highlighted, the rest grey" chart rendered in a single color. The preview now paints them, as PowerPoint does.
