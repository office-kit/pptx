---
'@office-kit/pptx': patch
---

fix: removing or replacing a link left its relationship behind on the slide

`setShapeClickAction(shape, null)`, `setShapeHyperlink(shape, null)` and `setShapeRunHyperlink(…, null)` deleted the `<a:hlinkClick>` but kept the `hyperlink` / `slide` relationship it pointed at, and replacing a link added a second one beside the first. The orphan stayed in `_rels`, so the URL remained in the saved file and a slide jump kept its target slide alive as a dependency: after removing that target, `duplicateSlide` failed with `Cannot duplicate missing dependency /ppt/slides/slideN.xml`.

Each of the three setters now releases a link relationship once nothing on the slide points at it. A relationship other runs still use is kept, and relationships of every other type are untouched.
