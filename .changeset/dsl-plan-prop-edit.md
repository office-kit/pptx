---
'@office-kit/pptx-dsl': minor
---

`@office-kit/pptx-dsl/source-edit` adds `planPropEdit`, which rewrites `x`, `y`, `width`, `height` and `rotation` on the JSX element that created a shape. A literal becomes the shortest decimal that reproduces the exact OOXML value (whole EMU, or 60000ths of a degree); any other expression keeps its meaning and gains a delta (`x={col * 2.5}` → `x={col * 2.5 + 0.3}`) that later edits fold into; a missing prop is inserted. It refuses spread attributes, non-shape elements and non-numeric values with a reason code, like `planTextEdit`.
