---
'@office-kit/pptx': minor
---

A bullet's automatic hanging indent now follows the paragraph's text size. `setShapeBulletStyle`, `setParagraphBullet`, the `bullets` option of `setShapeText` and the DSL's `bullets` / `bullet` used the same indent for every size — 0.375 in at the top level, which suits 18 pt text — so a 12–14 pt list put a wide gap between the bullet and its text. The indents are now scaled by the size of the paragraph's first run (or its end-of-paragraph size) over 18 pt: 1.5 times the size at the top level, so 14 pt text gets a 21 pt (0.29 in) hanging indent, and the deeper levels scale the same way. `setParagraphLevel` keeps moving an automatic indent to the new level at the same scale.

**Output change:** bulleted and numbered paragraphs whose text is not 18 pt are written with different `marL` / `indent` values than before. Paragraphs with no size of their own, and indents you set yourself, are unchanged.
