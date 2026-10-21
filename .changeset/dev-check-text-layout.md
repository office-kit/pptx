---
'@office-kit/pptx-dev': minor
---

Added `office-pptx check <deck.tsx> [--min-font-size 10]`. It builds the deck and measures every text box with the preview's layout engine: text that overflows its box is an error (exit code 1); text smaller than 10pt, a word split across lines, a wrapped paragraph ending on a line of one or two characters, and an overflow measured with an estimated width (a font the preview does not bundle) are warnings. Table and chart text is not measured. New projects from `office-pptx init` run it in `npm run check` after `tsc --noEmit`; in an existing project, add `&& office-pptx check deck.tsx` to the `check` script.
