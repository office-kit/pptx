---
'@office-kit/pptx-dsl': minor
---

feat: `Fill` accepts `autoFit`

`<Fill target={...} autoFit="normal">` sets text auto-fit on the placeholder it writes into, with the same values as `Text` (`'none'`, `'normal'`, `'shape'`). Text longer than a template placeholder can now be told to shrink without a `Raw` callback that selects the same shape a second time. Omitting `autoFit` keeps whatever the template's `<a:bodyPr>` says, as before.
