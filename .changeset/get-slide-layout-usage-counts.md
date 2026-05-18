---
'pptx-kit': minor
---

feat: `getSlideLayoutUsageCounts(pres)` — layout name → number-of-slides
histogram. Every layout enumerated by `getSlideLayouts` appears as a
key (count `0` for unreferenced layouts), so the function surfaces
unused layouts directly — useful for trimming template decks that
ship with placeholder layouts the working deck never picks up.
