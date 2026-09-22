---
"@office-kit/pptx": minor
"@office-kit/pptx-dev": patch
---

Add comment replies through `addSlideComment({ replyTo })` and `getCommentParent`, preserving PowerPoint p15 threading extensions through save and reload. Removing a comment also removes its descendant replies.

The preview comments dialog supports creating and editing replies in English and Japanese, shows each reply's parent, and deletes threads with undo/redo support.
