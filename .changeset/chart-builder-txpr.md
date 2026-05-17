---
'pptx-kit': minor
---

feat: chart builder writes back axis tick-label style + rotation via
`<c:txPr>`. New `axisTxPrElement(style, rotationDeg)` helper emits the
`<c:txPr><a:bodyPr rot/><a:lstStyle/><a:p><a:pPr><a:defRPr…/></a:pPr></a:p></c:txPr>`
payload from `categoryAxisLabelStyle` / `categoryAxisLabelRotationDeg`
and the value-axis counterparts. Closes the read/write gap for these
fields; round-trip test added.
