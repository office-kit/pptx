---
'@office-kit/pptx-dev': patch
---

Match PowerPoint's font offset editor by requiring a percent value and validating the supported range for superscript and subscript offsets.

Preserve mixed superscript and subscript offsets when an invalid entry is discarded and another font setting is applied.
