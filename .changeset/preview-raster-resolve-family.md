---
'@office-kit/pptx-preview': minor
---

Added `resolveFamily` to the options of `renderSlideToImage` / `renderSlideToRgba`, the same mapping `renderSlideToSvg` already takes for `textLayout: 'svg'`. The PNG path mapped every authored font name to a bundled substitute, so a caller that measured a deck's font with its own face (for example a `buildFontkitMeasurer` alias from the deck's font to a face in `fontFiles`) got lines wrapped for that face but painted in the substitute. Return a family loaded through `fontFiles` to draw and measure the run in it. East Asian characters still take the family of the first `fontFiles` face that covers kana; `resolveFamily` decides the rest of the run. Unset, the output is unchanged.
