---
"@office-kit/pptx-dev": patch
---

Fix Enter and Shift+Enter across an outline title and its body to split the slide and move the remaining body text into the new title. Preserve text links and restore the original slide with Undo.

Fix Enter and Shift+Enter across adjacent slide titles in the outline to retain both slides and their remaining text instead of inserting a line break into the first title.
