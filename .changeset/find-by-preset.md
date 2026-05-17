---
'pptx-kit': minor
---

feat: `findShapesByPreset(slide, preset)` returns every shape whose
`<a:prstGeom prst="…"/>` matches. Useful for diagram introspection:
find all `'leftArrow'`s for a workflow swap, replace every `'cloud'`
with `'rect'`, etc. Shapes without a preset (custGeom / pictures /
charts / tables / connectors / groups) are filtered out.
