---
'pptx-kit': minor
---

feat: round out gridline color round-trip with 3 more fields —
`valueAxisMinorGridlineColor`, `categoryAxisMajorGridlineColor`, and
`categoryAxisMinorGridlineColor`. Previously only the value-axis major
color was carried. All four now share a new chart-builder
`gridlinesElement(local, color?)` helper and a chart-reader
`readGridlineColor(gl)` helper; the existing major-gridline color
inline parse was replaced with a call to the shared reader for
consistency.
