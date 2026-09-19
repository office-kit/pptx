---
'@office-kit/pptx': minor
---

feat: author and read the paragraph-end format, and merge table cells without keeping the covered cells' text

- `setShapeParagraphs` / `setTableCellParagraphs` accept `endFormat` on each paragraph, written as `<a:endParaRPr>` after the runs. It is the place to give a paragraph with no runs a font size. Omitting `endFormat` writes nothing, as before.
- `getParagraphEndFormat(shape, paragraphIndex)` reads a paragraph's end-mark format (`null` when absent), and every paragraph returned by `getTableCellParagraphs` now carries `endFormat`.
- `mergeTableCells(table, block, { coveredText: 'drop' })` removes `<a:txBody>` from the cells the merge covers, the shape PptxGenJS writes. The default (`'keep'`) still leaves the covered cells' text in the XML. Covered cells without `<a:txBody>` read back as no paragraphs and can be written to again.
- `setTableCellText` / `setTableCellParagraphs` on a cell that has no `<a:txBody>` now always insert it as the cell's first child, so a cell that carries only `<a:extLst>` stays schema-valid.
- `setShapeParagraphs` / `setTableCellParagraphs` now validate every paragraph before they replace anything: a rejected format (for example a font size out of range) throws and leaves the existing text as it was.
