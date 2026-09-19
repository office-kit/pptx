// The preview must not trip over two table shapes the core can now author and
// that pptxgenjs always writes: a run-less paragraph that carries only an
// end-mark format (`<a:endParaRPr>`), and a merge-covered cell with no
// `<a:txBody>` at all.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTable,
  findSlideLayout,
  getSlides,
  getTableCell,
  inches,
  loadPresentation,
  mergeTableCells,
  savePresentation,
  setTableCellParagraphs,
} from '../src/api/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';
import { textContentOf } from './lib/svg-query.ts';

const fixture = (path: string): string =>
  fileURLToPath(new URL(`./fixtures/${path}`, import.meta.url));

const TEXT_LAYOUTS = ['svg', 'foreignObject'] as const;

describe('preview: end-mark-only paragraphs and txBody-less covered cells', () => {
  it.each(TEXT_LAYOUTS)('renders the pptxgenjs merged table (%s)', async (textLayout) => {
    const pres = await loadPresentation(await readFile(fixture('pptxgenjs/table-merge.pptx')));
    // pptxgenjs writes every row as `h="0"` (PowerPoint sizes it from the
    // text) and the preview does not auto-size rows, so the cells collapse to
    // zero height and no text is laid out. What this pins is that the covered
    // cells and the end-mark-only cell get through the cell walk: the anchors'
    // fills are painted and nothing throws.
    const svg = renderSlideToSvg(pres, getSlides(pres)[0]!, { textLayout });
    expect(svg.match(/fill="#F0F0F0"/g)).toHaveLength(3);
  });

  it.each(TEXT_LAYOUTS)('renders an authored table after a save (%s)', async (textLayout) => {
    const source = await loadPresentation(await readFile(fixture('minimal/blank.pptx')));
    const layout = findSlideLayout(source, 'Blank');
    if (!layout) throw new Error('Blank layout not found');
    const table = addSlideTable(addSlide(source, { layout }), {
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(1),
      rows: [
        ['', 'Group', ''],
        ['Row', '1', '2'],
      ],
    });
    setTableCellParagraphs(getTableCell(table, 0, 0), [{ runs: [], endFormat: { size: 9 } }]);
    mergeTableCells(table, { row: 0, col: 1, rowSpan: 1, colSpan: 2 }, { coveredText: 'drop' });

    const pres = await loadPresentation(await savePresentation(source));
    const svg = renderSlideToSvg(pres, getSlides(pres).at(-1)!, { textLayout });
    const text = textContentOf(svg);
    for (const visible of ['Group', 'Row', '1', '2']) expect(text).toContain(visible);
  });
});
