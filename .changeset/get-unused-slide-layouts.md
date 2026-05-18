---
'pptx-kit': minor
---

feat: `getUnusedSlideLayouts(pres)` — returns the layouts in the
package that no slide references. Useful when trimming a template
deck — unused layouts contribute parts and rels without ever
rendering. Iteration order matches `getSlideLayouts`.
