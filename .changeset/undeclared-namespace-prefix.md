---
'@office-kit/pptx': patch
---

Fixed: `duplicateSlide`, `addSlideImage` and other edits could save XML with an unbound `r:` (or other) prefix when the template declared that prefix on individual elements instead of the part root, so the saved file failed to load (#456). The writer now declares any prefix a new element or attribute uses that is not already in scope.
