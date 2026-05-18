---
'pptx-kit': minor
---

feat: `ChartSpec.valueAxisCrosses` — controls where the category axis
crosses the value axis. Accepts either an enum keyword
(`'autoZero' | 'min' | 'max'` → `<c:valAx><c:crosses val=…/>`) or a
numeric tagged form (`{ at: N }` → `<c:valAx><c:crossesAt val=N/>`).
The two forms are mutually exclusive per the schema; `crossesAt` wins
when both are present on read. Read by chart-reader, written by
chart-builder in the correct CT_ValAx schema order (after `<c:crossAx>`).
