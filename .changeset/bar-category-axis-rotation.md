---
'pptx-kit': minor
---

feat(site/playground): bar chart category-axis labels honor
`categoryAxisLabelRotationDeg`. The horizontal-value renderer used
the rotation field only for column charts (categories along the
x-axis); now the bar variant (categories down the y-axis) also
rotates each label around its anchor and widens its ellipsis budget
for tilted labels.
