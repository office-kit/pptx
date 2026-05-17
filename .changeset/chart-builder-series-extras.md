---
'pptx-kit': minor
---

feat: chart builder writes back series-level optional fields. Each
`<c:ser>` now emits:

- richer `<c:spPr>` with `<a:ln w="…"><a:prstDash/>` when
  `series.lineWidthEmu` or `lineDash` is authored
- `<c:invertIfNegative val="1"/>` when set
- `<c:marker><c:symbol/><c:size/></c:marker>` from `markerSymbol` /
  `markerSizePt`
- `<c:smooth val="1"/>` when set

Round-trip test covers all five fields.
