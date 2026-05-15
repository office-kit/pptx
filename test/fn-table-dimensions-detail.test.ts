// getTableColumnWidths / getTableRowHeights — richer counterpart to
// getTableDimensions (which only returns counts). Surfaces the EMU
// values stored on <a:tblGrid><a:gridCol w="..."> and <a:tr h="...">.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTable,
  emu,
  getSlides,
  getTableColumnWidths,
  getTableDimensions,
  getTableRowHeights,
  inches,
  insertTableColumn,
  insertTableRow,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getTableColumnWidths / getTableRowHeights', () => {
  it('reports column widths in EMU', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addSlideTable(slide, {
      x: inches(0), y: inches(0), w: inches(4), h: inches(2),
      rows: [['a', 'b'], ['c', 'd']],
      colWidths: [emu(914400), emu(914400 * 3)],
    });
    expect(getTableColumnWidths(table)).toEqual([914400, 914400 * 3]);
    expect(getTableDimensions(table)).toEqual({ rows: 2, cols: 2 });
  });

  it('reports row heights in EMU', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addSlideTable(slide, {
      x: inches(0), y: inches(0), w: inches(4), h: inches(2),
      rows: [['a', 'b'], ['c', 'd']],
      rowHeights: [emu(300000), emu(700000)],
    });
    expect(getTableRowHeights(table)).toEqual([300000, 700000]);
  });

  it('grows after insertTableRow / insertTableColumn', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addSlideTable(slide, {
      x: inches(0), y: inches(0), w: inches(4), h: inches(2),
      rows: [['a', 'b'], ['c', 'd']],
      colWidths: [emu(914400), emu(914400)],
    });
    expect(getTableColumnWidths(table)).toHaveLength(2);
    expect(getTableRowHeights(table)).toHaveLength(2);

    insertTableColumn(table);
    expect(getTableColumnWidths(table)).toHaveLength(3);
    expect(getTableRowHeights(table)).toHaveLength(2);

    insertTableRow(table);
    expect(getTableColumnWidths(table)).toHaveLength(3);
    expect(getTableRowHeights(table)).toHaveLength(3);
  });
});
