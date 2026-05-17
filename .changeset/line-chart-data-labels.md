---
'pptx-kit': minor
---

feat(site/playground): line and area charts paint per-point value labels
when `<c:dLbls><c:showVal val="1"/>` is set. Labels sit just above each
marker and route through the chart number-format projector (so
`<c:numFmt formatCode="0%"/>` etc. apply the same as on bar / pie).
Honors the per-series → chart-level cascade.
