---
'@office-kit/pptx-preview': minor
'@office-kit/pptx-dev': patch
---

Show bullet and numbered-list markers during inline editing, including nested numbering, without changing copied text or selection offsets. Share the preview renderer's numbering through `paragraphNumberLabels` so editing surfaces use the same counter and restart rules. Build inline paragraph text in one pass instead of copying the whole shape for every paragraph.
