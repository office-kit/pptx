---
'@office-kit/pptx-preview': patch
---

Themed Style 2 tables now show the same background gradient as the reference desktop app. The reference desktop app does not blend this two-stop table background at a constant rate: the top color holds through the header row, then the blend gets steeper toward the bottom. The preview used a constant rate, which made the middle rows up to 21 levels per channel too dark. It now uses the curve fitted to the reference desktop app's exports and matches them within 2 levels for all six accents.
