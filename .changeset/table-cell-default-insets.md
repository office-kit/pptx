---
'pptx-kit': patch
---

Emit PowerPoint's default cell insets on table cells

`addSlideTable` cells now carry the explicit default insets PowerPoint and
PptxGenJS both write — `<a:tcPr marL="91440" marR="91440" marT="45720"
marB="45720">` — plus a `<a:pPr marL="0" indent="0"><a:buNone/></a:pPr>` that
suppresses any inherited list bullet on the cell paragraph. The table renders
identically (these match the values PowerPoint applies when they're absent),
but the cell is now self-describing, so the output matches a PowerPoint- or
PptxGenJS-authored table byte-for-byte at the cell level.

Note: `getTableCellMargins` now returns the explicit `91440 / 45720` defaults
for a freshly-authored cell instead of `null`.
