---
'pptx-kit-site': patch
---

site(preview): keep text upright in flipped shapes. A shape's `flipH`/`flipV`
mirrors its geometry but text should stay readable (as in PowerPoint); the
preview was mirroring the text too. Text now follows the shape's rotation only,
not its flips.
