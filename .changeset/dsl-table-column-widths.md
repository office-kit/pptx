---
'@office-kit/pptx-dsl': minor
---

`Table` now throws when its `columnWidths` do not add up to its `width` (a difference of 0.01 in or more). The column widths decide how wide the table is drawn, so a mismatch was exported silently and the table overflowed or fell short of the space laid out for it.
