---
'@office-kit/pptx-dsl': minor
---

feat: `bullet` and `level` on single `Text` paragraphs

`bullets` and `paragraphSpacing` (0.4.0) apply to every paragraph. This adds the per-paragraph half, for nested lists and for a heading line that stays out of the list.

- Each entry of `paragraphs` accepts `bullet` (`'bullet'`, `'number'`, `'none'`, `{ char }`, `{ autoNum }`) and `level` next to the core `ParagraphSpec` fields. The new exported types are `TextParagraph` and `TextLevel`.
- A paragraph's `bullet` wins over `bullets`, so `bullet: 'none'` exempts one line.
- `level` is typed `0` to `8`, so `tsc` rejects a level outside that range. A nested bullet is indented deeper than its parent.
- A nested list no longer needs a `Raw` callback that addresses paragraphs by index.
