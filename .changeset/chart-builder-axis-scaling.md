---
'pptx-kit': minor
---

feat: chart builder writes back full value-axis scaling. `<c:valAx>`
now emits `<c:scaling><c:min/>/<c:max/>`, `<c:numFmt formatCode>`,
`<c:majorUnit>`, and `<c:minorUnit>` when authored — completing the
read/write parity for `ChartSpec.valueAxis`. Round-trip test added.
