---
'@office-kit/pptx-preview': patch
---

Keep long Latin words intact in SVG previews unless the paragraph explicitly allows mid-word wrapping, matching PowerPoint's default for the DrawingML `latinLnBrk` setting.

When mid-word wrapping is enabled, use the remaining line width while keeping East Asian closing punctuation attached to its preceding character.
