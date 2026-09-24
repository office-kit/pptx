// Table row + column insert/remove.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTable,
  getSlides,
  getTableCellText,
  getTableCellSpan,
  mergeTableCells,
  savePresentation,
  getSlideTables,
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

for (const rows of [true, false]) {
  for (const index of [0, 1, 2]) {
    it(`inserting ${rows ? 'row' : 'column'} at ${index} preserves a 2×2 merge`, async () => {
      const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
      const table = addBaseTable(getSlides(pres)[0]!);
      mergeTableCells(table, { row: 0, col: 0, rowSpan: 2, colSpan: 2 });
      if (rows) insertTableRow(table, index);
      else insertTableColumn(table, index);
      const check = (table: ReturnType<typeof addBaseTable>) => {
        const cells = getTableCells(table);
        const startRow = rows && index === 0 ? 1 : 0;
        const startCol = !rows && index === 0 ? 1 : 0;
        const rowSpan = rows && index === 1 ? 3 : 2;
        const colSpan = !rows && index === 1 ? 3 : 2;
        expect(getTableCellSpan(cells[startRow]![startCol]!)).toEqual({
          rowSpan,
          gridSpan: colSpan,
          hMerge: false,
          vMerge: false,
        });
        expect(getTableCellText(cells[startRow]![startCol]!)).toBe('A');
        for (let r = 0; r < cells.length; r++)
          for (let c = 0; c < cells[r]!.length; c++) {
            if (r === startRow && c === startCol) continue;
            const inside =
              r >= startRow && r < startRow + rowSpan && c >= startCol && c < startCol + colSpan;
            expect(getTableCellSpan(cells[r]![c]!)).toEqual({
              rowSpan: 1,
              gridSpan: 1,
              hMerge: inside && c > startCol,
              vMerge: inside && r > startRow,
            });
          }
      };
      check(table);
      const loaded = await loadPresentation(await savePresentation(pres));
      check(getSlideTables(getSlides(loaded)[0]!)[0]!);
    });
  }
}

for (const rows of [true, false]) {
  for (const index of [0, 1]) {
    it(`deleting ${rows ? 'row' : 'column'} ${index} shrinks a merge and preserves its anchor`, async () => {
      const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
      const table = addBaseTable(getSlides(pres)[0]!);
      mergeTableCells(table, { row: 0, col: 0, rowSpan: 2, colSpan: 2 });
      const remove = rows ? removeTableRow : removeTableColumn;
      for (const invalid of [NaN, 0.5, -1, 2])
        expect(() => remove(table, invalid)).toThrow(RangeError);
      remove(table, index);
      const check = (table: ReturnType<typeof addBaseTable>) => {
        const cells = getTableCells(table);
        expect(cells).toHaveLength(rows ? 1 : 2);
        expect(cells[0]).toHaveLength(rows ? 2 : 1);
        expect(getTableCellText(cells[0]![0]!)).toBe('A');
        expect(getTableCellSpan(cells[0]![0]!)).toEqual({
          rowSpan: rows ? 1 : 2,
          gridSpan: rows ? 2 : 1,
          hMerge: false,
          vMerge: false,
        });
        const covered = rows ? cells[0]![1]! : cells[1]![0]!;
        expect(getTableCellSpan(covered)).toEqual({
          rowSpan: 1,
          gridSpan: 1,
          hMerge: rows,
          vMerge: !rows,
        });
      };
      check(table);
      check(
        getSlideTables(getSlides(await loadPresentation(await savePresentation(pres)))[0]!)[0]!,
      );
    });
  }
}
