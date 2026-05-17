// setTableColumnWidth / setTableRowHeight — resize a single column or
// row by index. Counterpart to getTableColumnWidths / getTableRowHeights.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTable,
  emu,
  getSlides,
  getTableColumnWidths,
  getTableRowHeights,
  inches,
  loadPresentation,
  setTableColumnWidth,
  setTableRowHeight,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: setTableColumnWidth / setTableRowHeight', () => {
  it('resizes a single column without affecting others', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addSlideTable(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(6),
      h: inches(2),
      rows: [
        ['a', 'b', 'c'],
        ['d', 'e', 'f'],
      ],
      colWidths: [emu(1000000), emu(1000000), emu(1000000)],
    });
    setTableColumnWidth(table, 1, emu(2500000));
    expect(getTableColumnWidths(table)).toEqual([1000000, 2500000, 1000000]);
  });

  it('resizes a single row without affecting others', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addSlideTable(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      rows: [
        ['a', 'b'],
        ['c', 'd'],
      ],
      rowHeights: [emu(500000), emu(500000)],
    });
    setTableRowHeight(table, 0, emu(900000));
    expect(getTableRowHeights(table)).toEqual([900000, 500000]);
  });

  it('throws on out-of-range column / row indices', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addSlideTable(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      rows: [['a']],
    });
    expect(() => setTableColumnWidth(table, 9, emu(1))).toThrow(/out of range/);
    expect(() => setTableRowHeight(table, 9, emu(1))).toThrow(/out of range/);
  });
});
