---
"@office-kit/pptx-dev": patch
---

Add bilingual corner handles for proportionally resizing multiple selected objects. Scale positions and dimensions together while preserving rotations and the opposite corner. Keep each resize in one Undo/Redo step and support Escape cancellation. This changes object geometry; font sizes and other appearance attributes retain their existing values.
