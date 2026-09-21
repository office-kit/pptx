---
'@office-kit/pptx': minor
---

Placeholder lookups and slide transitions now reject an unknown token at compile time instead of at runtime.

`findSlidePlaceholder`, `findSlidePlaceholders` and `findLayoutsWithPlaceholderType` take the closed `ST_PlaceholderType` set (exported as `PlaceholderType`) rather than `string`, so a shape's display name or a typo no longer type-checks and then silently finds nothing. `setSlidePlaceholders` keys its `byType` record the same way, so a misspelled key is no longer skipped in silence.

`TransitionEffect` gains the five spec effects it was missing — `comb`, `pull`, `random`, `strips` and `wheel` — and `TransitionOptions.effect` drops its `| string` escape, so a mistyped effect is a type error rather than a runtime throw. Reading stays permissive: `getSlideTransition` returns the new `SlideTransition` type, whose `effect` is still `string`, because a deck authored elsewhere can carry an effect this library does not model.
