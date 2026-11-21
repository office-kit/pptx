---
'@office-kit/pptx': minor
---

`setShapeParagraphs` and `setTableCellParagraphs` accept an equation as a run: `{ equation, display?, format? }`, where `equation` is a tree of `MathNodeSpec` nodes (the input counterpart of `MathNode`). It is written as a native `<a14:m>` equation inside `mc:AlternateContent`, with the equation's linear text (`σ(z)=1/(1+e^(−z))`) as the fallback run that apps without equation support show. Equations in that form now also read back through `getShapeParagraphElements`. `ParagraphSpec.runs` is now `ReadonlyArray<RunSpec | EquationSpec>`, so code that reads `run.text` from a spec needs to check for `equation` first.
