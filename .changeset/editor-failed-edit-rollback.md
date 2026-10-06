---
"@office-kit/pptx-dev": patch
---

A canvas edit that fails partway is now undone completely instead of leaving
half-applied changes (such as an unused image) in the saved deck. Project Undo
and Redo no longer leave a source file partially written when a write fails,
and a refused text edit keeps its original reason when the source also changed.

Undo and Redo now replace a project file by writing a temporary file beside it
and renaming it, so a file inside a read-only directory can no longer be
restored in place.
