---
'@office-kit/pptx-dev': minor
---

Added `office-pptx check <deck.tsx> [--min-font-size 10]`. It builds the deck and measures every text box with the preview's layout engine: text that overflows its box or is smaller than 10pt is an error (exit code 1), and a word split across lines or a wrapped paragraph ending on a line of one or two characters is a warning. Table and chart text is not measured. New projects from `office-pptx init` run it in `npm run check` after `tsc --noEmit`; in an existing project, add `&& office-pptx check deck.tsx` to the `check` script.
