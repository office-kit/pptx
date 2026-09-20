---
'@office-kit/pptx-preview': patch
---

Keep a resized group's text at its authored size in the preview. A group whose `<a:ext>` differs from its `<a:chExt>` scales its children, and the renderer was applying that scale to their glyphs as well, so a group squashed vertically (what Google Slides writes for a hand-resized group) drew stretched, half-height letters. PowerPoint and LibreOffice resize only the geometry, so the text now lays out inside the group-scaled rect at its authored point size and aspect.
