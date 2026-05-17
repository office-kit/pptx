---
'pptx-kit': minor
---

feat(site/playground): include slide-master count in the
"masters · layouts · sections" meta cell. `getPresentationSummary`
already returned layout / section counts; the playground now also
calls `getSlideMasterCount` so multi-master decks surface that fact
in the audit panel.
