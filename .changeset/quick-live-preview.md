---
"@office-kit/pptx-dev": patch
---

Keep slide previews steady while editing TSX: update only changed thumbnails and slides, preserve zoom and scroll position, and retain the previous frame until its replacement is ready. Transfer SVG changes incrementally, preload build workers between saves, and cancel obsolete evaluations so rapid edits and accidental infinite loops do not delay the next revision.
