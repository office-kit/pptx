---
'pptx-kit': patch
---

Fix jammed bullet lists and unreadable chart categories

- **Bulleted text boxes now indent correctly.** `setShapeBullets` /
  `setParagraphBullet` added the bullet glyph but no hanging indent, so a bullet
  authored on a text box (which inherits the master's `otherStyle`, marL=0, not
  the body style) rendered with the glyph jammed against the text. They now
  write PowerPoint's per-level default `marL` / `indent` (unless the caller set
  their own), matching PowerPoint and PptxGenJS.
- **Charts with multi-level category references now read back.** The chart
  reader handled `<c:strRef>` / `<c:strLit>` categories but not
  `<c:multiLvlStrRef>`, which is what PowerPoint and PptxGenJS emit — so
  `getShapeChartCategories` (and the full `getShapeChartSpec`) returned an empty
  category list for those charts. It now reads the level's points.
