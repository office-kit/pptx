// Table row + column insert/remove.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTable,
  getSlides,
  getTableCellText,
  getTableCells,
  inches,
  insertTableColumn,
  insertTableRow,
  loadPresentation,
  removeTableColumn,
  removeTableRow,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const addBaseTable = (slide: ReturnType<typeof getSlides>[number]) =>
  addSlideTable(slide, {
    x: inches(0),
    y: inches(0),
    w: inches(4),
    h: inches(2),
    rows: [
      ['A', 'B'],
      ['C', 'D'],
    ],
  });

describe('fn API: table row/column mutation', () => {
  it('insertTableRow appends at end by default', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addBaseTable(slide);
    insertTableRow(table, undefined, ['E', 'F']);
    const cells = getTableCells(table);
    expect(cells).toHaveLength(3);
    expect(getTableCellText(cells[2]![0]!)).toBe('E');
    expect(getTableCellText(cells[2]![1]!)).toBe('F');
  });

  it('insertTableRow at index 0 prepends', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addBaseTable(slide);
    insertTableRow(table, 0, ['header-A', 'header-B']);
    const cells = getTableCells(table);
    expect(getTableCellText(cells[0]![0]!)).toBe('header-A');
    expect(getTableCellText(cells[1]![0]!)).toBe('A');
  });

  it('removeTableRow drops the row', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addBaseTable(slide);
    removeTableRow(table, 0);
    const cells = getTableCells(table);
    expect(cells).toHaveLength(1);
    expect(getTableCellText(cells[0]![0]!)).toBe('C');
    expect(() => removeTableRow(table, 99)).toThrow(RangeError);
  });

  it('insertTableColumn adds a new column at the end', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addBaseTable(slide);
    insertTableColumn(table);
    const cells = getTableCells(table);
    expect(cells[0]).toHaveLength(3);
    expect(getTableCellText(cells[0]![2]!)).toBe('');
  });

  it('insertTableColumn at index inserts in the middle', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addBaseTable(slide);
    insertTableColumn(table, 1, 500000);
    const cells = getTableCells(table);
    expect(cells[0]).toHaveLength(3);
    // Original 'B' shifted to index 2.
    expect(getTableCellText(cells[0]![1]!)).toBe('');
    expect(getTableCellText(cells[0]![2]!)).toBe('B');
  });

  it('removeTableColumn drops the column from every row', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addBaseTable(slide);
    removeTableColumn(table, 0);
    const cells = getTableCells(table);
    expect(cells[0]).toHaveLength(1);
    expect(getTableCellText(cells[0]![0]!)).toBe('B');
    expect(() => removeTableColumn(table, 99)).toThrow(RangeError);
  });
});
