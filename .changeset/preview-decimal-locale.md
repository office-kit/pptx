---
'@office-kit/pptx-preview': minor
---

Decimal tab stops align on the decimal separator of each run's language (`,` for `de-DE`, `fr-FR`, …) instead of always `.`. Runs without a language still use `.`. `decimalSeparatorOf(lang)` is exported. Paragraph indents in vertical text rendered as HTML now run along the text lines, as in PowerPoint, instead of being applied as a left margin.
