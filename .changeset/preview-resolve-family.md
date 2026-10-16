---
'@office-kit/pptx-preview': minor
---

Added `resolveFamily` to `RenderSlideOptions` (#461). With `textLayout: 'svg'`, every font other than the Calibri / Cambria / Arial / Times / Courier families used to be written to the SVG as `Carlito`, so a host that embeds its own `@font-face` rules (for example for a deck that uses Lora or Inter) could not get its font applied. `resolveFamily` maps each authored font name (or `null` when none applies) to the family written to `<text font-family>` and passed to `measureText`, and the shrink-to-fit search measures with it too. It defaults to `substituteFamily`, so output is unchanged unless you pass it. The `foreignObject` text path already keeps the authored name and ignores the option.
