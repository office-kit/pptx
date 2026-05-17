---
'pptx-kit': minor
---

feat(chart): `ChartSpec.dispBlanksAs` reads `<c:dispBlanksAs>`
(`'gap' | 'zero' | 'span'`). Playground line / area renderer:

- `gap` (default): breaks the path on null values
- `zero`: substitutes 0 so the line dips to the baseline
- `span`: connects the surrounding points across the gap

Previously every null value was coerced to 0, which silently
flattened the chart whenever the deck had genuine missing data.
