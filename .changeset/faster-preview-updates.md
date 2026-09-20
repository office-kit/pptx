---
"@office-kit/pptx-dev": patch
---

Reduce TSX save-to-preview latency by reusing the compiler and unchanged slide renders, while preserving fresh deck evaluation and invalidating previews when shared resources change.
