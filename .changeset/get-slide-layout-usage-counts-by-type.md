---
'pptx-kit': minor
---

feat: `getSlideLayoutUsageCountsByType(pres)` — companion to
`getSlideLayoutUsageCounts`, but keyed on the OOXML layout-type enum
token (`title`, `obj`, `twoObj`, `blank`, …) instead of the user-
visible name. Stable across PowerPoint UI locales. Useful for "how
many content slides vs. dividers vs. title slides?" audits.
