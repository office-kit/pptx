---
'@office-kit/pptx': minor
---

Colors are now a type, not a string. `Color` (`#RRGGBB`, the `#RGB` shorthand, a theme token such as `accent1`, or the `scheme:`-prefixed form) replaces `string` on every authoring call that takes one — shape and table fills, slide backgrounds, strokes, shadows, glows, gradient stops, connectors and run formats — so `'reddish'` or a mistyped `'accnet1'` is a compile error instead of a run-time throw. Chart colors take `HexColor`, which enforces what the docs already said: a chart series must resolve to a concrete sRGB value and cannot carry a theme token.

The `#` is now required. Bare `RRGGBB` was one of four accepted spellings of the same thing, and admitting it collapsed the union back to `string`, taking every other color check down with it. The runtime parser still accepts it, so untyped JavaScript callers are unaffected.

`@office-kit/pptx-dsl` picks this up too: `fill`, `stroke`, `background`, `stripeFill` and cell borders are typed the same way, so a bad color in a TSX deck fails `npm run check` rather than the export.

Reading stays permissive, since a deck authored elsewhere can carry a scheme token outside its theme. `getShapeRunFormat` and friends now return `ReadTextFormat`, and the gradient readers return `ReadGradientFill` / `ReadGradientStop` — same shapes, with colors as `string`. To write a color you read back, pass it through the new `asColor`, which returns `null` when the value is not one this library can emit.
