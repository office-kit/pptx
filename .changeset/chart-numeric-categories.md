---
'pptx-kit': minor
---

fix: chart categories accept `<c:cat><c:numRef>` (numeric / date
categories). Previously the category-axis dropped to an empty
labels array when the chart authored a numeric category channel
(common for date-axis line charts authored in Excel). Falls back
to formatting each cached numeric value as a string so date /
number cats appear on the axis instead of disappearing.
