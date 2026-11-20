---
'@office-kit/pptx': minor
'@office-kit/pptx-preview': minor
'@office-kit/pptx-editor': patch
'@office-kit/pptx-dev': patch
---

Equations (`<a14:m>`) in a paragraph now read as `math` elements of `getShapeParagraphElements` and `getTableCellParagraphs`, one per `<m:oMath>`. Each carries a typed `MathNode` tree (fractions, scripts, radicals, n-ary operators, delimiters, matrices, accents, functions and limits) plus its plain run text; other OMML elements read as `unsupported` nodes with their text. Code that switches on `kind` should handle the new `'math'` member; a `math` element's `format` is always `null`.

The preview draws these equations in place of the blank it left before, in both text layouts, with display equations (`<m:oMathPara>`) at full size. Elements outside the modeled set are drawn as plain text in a dashed box and reported by `auditTextLayout` as `unsupported-math`. The editor keeps ignoring equations when it counts text offsets, as the core text-range functions do.
