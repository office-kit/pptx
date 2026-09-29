---
'@office-kit/pptx-dsl': minor
---

feat: trace a compiled shape back to the TSX that created it, and edit literal text without a filesystem

- `getShapeJsxSources(shape)` returns the source locations (1-based line and column, outermost first) of the elements that were being evaluated when `compile()` created the shape, when the deck was built with the dev JSX transform. A shape made by a component, including a prebuilt one, carries the component's call site.
- A component that returns an element with its own source location now keeps that location under its call site, so compile errors name both.
- New `@office-kit/pptx-dsl/source-edit` subpath: `planTextEdit` and `verifyTextEdit` apply the dev preview's literal text-edit rules to source strings and report failures as reason codes (`invalid`, `not-unique-on-slide`, `not-found`, `ambiguous`, `slide-count-changed`, `other-slides-changed`, `text-mismatch`). An `anchor` from `getShapeJsxSources` narrows the search to one element. `typescript` is an optional peer dependency needed only by this subpath.
