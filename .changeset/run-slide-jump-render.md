---
'pptx-kit': minor
---

feat(site/playground): per-run slide-jump click actions render as
in-page anchors. Mirrors the shape-level slide-jump support shipped
in the prior batch — `getShapeRunClickAction` resolves to either a
URL or `#slide-N` anchor, and the run-level `<a href>` wrapper
respects whether the href is in-page (no `target=_blank`) or
external.
