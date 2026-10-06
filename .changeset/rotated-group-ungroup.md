---
"@office-kit/pptx": patch
"@office-kit/pptx-dev": patch
---

Preserve group rotation and reflections when ungrouping objects. Children retain their transformed positions and compose the group angle and flips with their own, including nested groups. The development preview supports this through the existing English and Japanese ungroup controls and Undo/Redo.
