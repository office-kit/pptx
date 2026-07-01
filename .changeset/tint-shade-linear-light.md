---
'pptx-kit': patch
---

Fix `<a:tint>` / `<a:shade>` colour resolution to compute in linear-light RGB,
matching PowerPoint and LibreOffice. A 75% tint of black now resolves to a mid
grey (~#8B8B8B) instead of the too-dark #404040, so colours derived from theme
scheme transforms (subtitle placeholders, table banding, chart fills) render at
the right lightness.
