---
'@office-kit/pptx': minor
---

Add `getPresentationTextReview(pres, { mode, full })`, which returns the deck review the skill's `deck-text.mjs` prints: the headline track, `[要確認]` markers left on slides and in notes, Japanese wording flags, and with `mode: 'talk' | 'document'` the matching headline checks. Servers and workers can now produce that review from a loaded presentation without running the script. The script now calls this function, so it needs `@office-kit/pptx` 0.27.0 or later in the slide project and says so when the installed copy is older.
