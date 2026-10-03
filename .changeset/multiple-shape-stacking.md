---
'@office-kit/pptx': minor
'@office-kit/pptx-dev': patch
---

Move all selected objects together with Bring to Front, Send to Back, Bring Forward and Send Backward, preserving their relative stacking order and supporting a single undo step in the editor.

The four existing stacking APIs also accept an array of sibling shapes. Selections from different parent containers are rejected before any shape is moved.
