---
'pptx-kit': minor
---

feat: `ChartTrendline.name` — round-trip a custom trendline label
(`<c:trendline><c:name>…`). PowerPoint auto-generates a label like
"Linear (X)" or "MA(5) (X)" when this element is omitted; authors who
want a different label (or who imported one from another tool) now
have the field. Read by chart-reader; written by chart-builder at the
CT_Trendline schema-required first position (before `<c:spPr>`).
