---
'pptx-kit': minor
---

feat: chart value-axis honors `<c:dispUnits><c:builtInUnit/>`.
`ChartAxisScaling.displayUnits` carries the authored scale token
(`hundreds`, `thousands`, `millions`, etc.). The playground divides
each value-axis tick by the corresponding divisor before formatting,
so charts authored "in millions" finally render as `10` / `20` /
`30` instead of `10000000`.
