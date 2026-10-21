---
'@office-kit/pptx-dev': minor
---

Added `office-pptx check <deck.tsx> [--min-font-size 10]`. It builds the deck and measures every text box with the preview's layout engine: text that runs out of its shape is an error (exit code 1); text that only spills into the box's inner margins, text smaller than 10pt, a word split across lines, and a wrapped paragraph ending on a line of one or two characters are warnings. When a width had to be estimated (a font the preview does not bundle, such as any Japanese font), only text that runs more than a line's height out of its shape (above and below a centred box counted together, or past either side) is an error; smaller overflows are warnings. Table and chart text is not measured. New projects from `office-pptx init` run it in `npm run check` after `tsc --noEmit`; in an existing project, add `&& office-pptx check deck.tsx` to the `check` script.
