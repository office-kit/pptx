---
'@office-kit/pptx': minor
'@office-kit/pptx-dev': patch
---

Persist outline collapse with `getCollapsedOutlineSlides` and `setSlideOutlineCollapsed`, preserving PowerPoint's per-slide view state. Deleting a slide removes its outline reference.

In the editor, double-click a slide icon in Outline View to collapse or expand its body. The change supports Undo and survives saving and reopening.
