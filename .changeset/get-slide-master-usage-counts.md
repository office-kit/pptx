---
'pptx-kit': minor
---

feat: `getSlideMasterUsageCounts(pres)` — master part name → number of
slides chaining to that master. Every master in the package appears as
a key (count `0` for unreferenced masters), so it surfaces unused
masters directly. Pair with `getSlideLayoutUsageCounts` for the
layout layer in multi-master template decks.
