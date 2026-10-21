---
'@office-kit/pptx-preview': patch
---

The preview now links a deck's hyperlinks only when they point to `http:`, `https:`, `mailto:` or another slide, so a `javascript:`, `data:` or `vbscript:` link can no longer run script in the page that shows the SVG, and font names are quoted in inline styles so they cannot add CSS declarations.
