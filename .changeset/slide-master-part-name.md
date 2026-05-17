---
'pptx-kit': minor
---

feat: `getSlideMasterPartName(slide)` returns the part-name of the
slide master the slide inherits from. Useful for multi-master decks
where different slides live under different brand templates and the
caller needs to scope theme / fontScheme / clrMap lookups to the
correct master.
