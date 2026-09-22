---
"@office-kit/pptx": minor
---

`getSlideAnimations` reads what a timing tree says becomes of an effect's value once it has run. Each step reports it as `valueAfterEnd`: `'held'` when every behaviour that animates something carries a `fill` that keeps its value, `'removed'` when one says it is taken away again, and `'unstated'` when the tree does not say — the schema gives `fill` no default, so an absent one is not the same as `'hold'`.

A step is now `playable` (and so `editable`) only when the tree states that the effect leaves the shape where it puts it: the effect node's own `fill`, and that of the `<p:set>` that flips visibility — or, for an effect that only fades, the fade's. An effect that says its value is taken away there, or says nothing there, is still listed and still saved exactly as it arrived, but is reported as one this library reads rather than plays.

`valueAfterEnd` itself decides nothing: an effect whose fade is removed or unstated still plays, because on its own the shape ends up where its own visibility puts it. It is what a player needs when two effects animate one object at the same time and the shorter one ends first — whether its value stays above the longer one or gives way to it. Decks this library authored are unaffected: it has always written `fill="hold"`.
