---
'@office-kit/pptx-dev': patch
---

`npm install @office-kit/pptx-dev` failed with a 404 because every release since 0.12.0 depends on `@office-kit/pptx-editor`, which is not on npm yet. pptx-dev now bundles the editor code it runs, so it no longer depends on that package.
