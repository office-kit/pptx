---
'pptx-kit': minor
---

feat(site/playground): each shape's authored name surfaces as a
`data-pptx-shape-name` attribute on its root `<g>` element. Lets
DevTools, a11y inspectors, or test selectors target shapes by their
PowerPoint name without parsing SVG geometry. Cheap to emit and has
no visual impact.
