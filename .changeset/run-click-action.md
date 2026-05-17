---
'pptx-kit': minor
---

feat: `getShapeRunClickAction(shape, p, r)` returns the per-run
hlinkClick action with the same `ShapeClickAction` discriminated
union the shape-level `getShapeClickAction` uses. Recognises external
URLs, slide-jump (`ppaction://hlinksldjump`), and the four
nav-preset actions (next / prev / first / last slide). Lets callers
treat per-run hyperlinks symmetrically with shape-level ones.
