---
'@office-kit/pptx-editor': minor
'@office-kit/pptx-dev': minor
---

Outline View now edits across slide boundaries as Mac PowerPoint's outline does:

- **Select across slides** by dragging with the mouse or Shift-clicking, as well as with Shift+arrow keys.
- **Delete, Cut, typing, pasting and Enter** over a selection that crosses a slide title remove the slides whose titles are selected: the remaining text after the selection joins the paragraph where it starts, and the last slide's remaining body moves up. This now works from any title or body paragraph, not only from a title. Slides with other objects ask for confirmation first.
- **Backspace** at the start of a slide title (or **Delete** at the end of the text before it) merges that slide into the previous one.
- **Drag a bullet** to move the paragraph and its sub-points to any position, including another slide; dragging sideways changes their level, and dragging a top-level bullet left turns it into a new slide.
- **Drag a slide icon to the right** to demote the slide into the previous slide's body.
- **Drag selected text** to move it, or hold Option (Control elsewhere) to copy it, including into another slide.

Each of these is one Undo step.
