import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import {
  addSlideTable,
  getSlides,
  getTableCellPosition,
  getTableCells,
  inches,
  loadPresentation,
  mergeTableCells,
} from '../src/api/index.ts';
import {
  neighboringTableCell,
  tableCellsInRange,
  tableSelectionBlock,
} from '../site/src/lib/editor/core/table-selection.ts';

it('moves in all directions through merged anchors and stops at table edges', async () => {
  const pres = await loadPresentation(
    await readFile(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url)),
  );
  const table = addSlideTable(getSlides(pres)[0]!, {
    x: inches(0),
    y: inches(0),
    w: inches(6),
    h: inches(3),
    rows: [
      ['A', 'B', 'C'],
      ['D', 'E', 'F'],
      ['G', 'H', 'I'],
    ],
  });
  mergeTableCells(table, { row: 0, col: 0, rowSpan: 2, colSpan: 2 });
  const block = tableSelectionBlock({ row: 2, col: 2, end: { row: 1, col: 1 } });
  expect(block).toEqual({ row: 1, col: 1, rowSpan: 2, colSpan: 2 });
  expect([...tableCellsInRange(getTableCells(table), block)].map(getTableCellPosition)).toEqual([
    { row: 0, col: 0 },
    { row: 0, col: 1 },
    { row: 1, col: 0 },
    { row: 1, col: 1 },
    { row: 1, col: 2 },
    { row: 2, col: 1 },
    { row: 2, col: 2 },
  ]);
  expect(neighboringTableCell(table, 0, 0, 0, 1)).toEqual({ row: 0, col: 2 });
  expect(neighboringTableCell(table, 0, 0, 1, 0)).toEqual({ row: 2, col: 0 });
  expect(neighboringTableCell(table, 1, 2, 0, -1)).toEqual({ row: 0, col: 0 });
  expect(neighboringTableCell(table, 2, 1, -1, 0)).toEqual({ row: 0, col: 0 });
  expect(neighboringTableCell(table, 0, 0, -1, 0)).toBeNull();
  expect(neighboringTableCell(table, 0, 0, 0, -1)).toBeNull();
  expect(neighboringTableCell(table, 2, 2, 1, 0)).toBeNull();
  expect(neighboringTableCell(table, 2, 2, 0, 1)).toBeNull();
});
