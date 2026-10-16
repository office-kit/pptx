---
'@office-kit/pptx-preview': minor
---

Added `resolveFamily` to `RenderSlideOptions` (#461). With `textLayout: 'svg'`, each run's font is written to the SVG as one of the bundled substitutes (Calibri / Aptos and unknown fonts as `Carlito`, Cambria as `Caladea`, Arial / Helvetica as `Liberation Sans`, Times as `Liberation Serif`, Courier / Consolas as `Liberation Mono`), so a host that embeds its own `@font-face` rules — for a deck that uses Lora or Inter, say — could not get its font applied. `resolveFamily` maps each authored font name (or `null` when none applies) to the family written to `<text font-family>` and passed to `measureText`; the shrink-to-fit search measures with it too, and names that are not plain CSS identifiers (such as `M PLUS 1p`) are quoted. It defaults to `substituteFamily`, so output is unchanged unless you pass it. The `foreignObject` text path already keeps the authored name and ignores the option.

Fixed: `auditTextLayout` sized a bare `<a:normAutofit/>` body's shrink with the substitute font while laying it out with the authored one, so with a `measureText` that registers that font the audit could report overflow (or miss it) for text the shrink was sized for.
