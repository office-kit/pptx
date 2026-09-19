---
'@office-kit/pptx': minor
---

feat: author and read the paragraph-end format, and merge table cells without keeping the covered cells' text

- `setShapeParagraphs` / `setTableCellParagraphs` accept `endFormat` on each paragraph, written as `<a:endParaRPr>` after the runs. A paragraph with no runs takes its line height from the end mark's `size`; without it PowerPoint falls back to the inherited default (18 pt), which stretches an empty table row. Omitting `endFormat` writes nothing, as before.
- `getParagraphEndFormat(shape, paragraphIndex)` reads a paragraph's end-mark format (`null` when absent), and every paragraph returned by `getTableCellParagraphs` now carries `endFormat`.
- `mergeTableCells(table, block, { coveredText: 'drop' })` removes `<a:txBody>` from the cells the merge covers, the shape PptxGenJS writes. The default still keeps the covered cells' text so a merge stays reversible. Covered cells without `<a:txBody>` read back as no paragraphs and can be written to again.
