---
'@office-kit/pptx': minor
'@office-kit/pptx-preview': patch
'@office-kit/pptx-dev': patch
---

Resolve inherited text autofit and columns through `getShapeBodyPrEffective`, and use them consistently in previews and editing controls. Allow `setShapeTextColumns` to author one column explicitly, overriding inherited columns; invalid column settings leave the previous values intact.
