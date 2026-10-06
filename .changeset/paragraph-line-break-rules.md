---
'@office-kit/pptx-preview': patch
---

Keep ordinary Latin words intact in SVG previews unless the paragraph explicitly allows mid-word wrapping, while still emergency-splitting words that cannot fit on an empty line as PowerPoint does. This matches the DrawingML `latinLnBrk` setting and PowerPoint's handling of narrow preset shapes.

When mid-word wrapping is enabled, use the remaining line width while keeping East Asian closing punctuation attached to its preceding character.
