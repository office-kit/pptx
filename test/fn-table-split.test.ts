import { unzipSync, strFromU8 } from 'fflate';
import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTable,
  createPresentation,
  getSlides,
  getSlideTables,
  getTableCell,
  getTableCellParagraphs,
  getTableCellSpan,
  getTableCellText,
  getTableColumnWidths,
  getTableRowHeights,
  getShapeBounds,
  inches,
  loadPresentation,
  mergeTableCells,
  savePresentation,
  setTableCellParagraphs,
  splitTableCell,
} from '../src/api/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

function fixture() {
  const p = createPresentation();
  const table = addSlideTable(addBlankSlide(p), {
    rows: [
      ['A', 'B'],
      ['C', 'D'],
    ],
    x: inches(1),
    y: inches(1),
    w: inches(6),
    h: inches(2),
  });
  setTableCellParagraphs(getTableCell(table, 0, 0), [
    { align: 'right', runs: [{ text: '日本語🙂', format: { bold: true, color: '#FF0000' } }] },
  ]);
  return { p, table };
}

describe('splitTableCell', () => {
  it('subdivides a normal cell and preserves its rich text and neighbor geometry on reload', async () => {
    const { p, table } = fixture();
    const before = getTableCellParagraphs(getTableCell(table, 0, 0));
    const bounds = getShapeBounds(table);
    splitTableCell(table, 0, 0, { rows: 2, columns: 3 });
    const loaded = await loadPresentation(await savePresentation(p));
    const again = getSlideTables(getSlides(loaded)[0]!)[0]!;
    expect(getShapeBounds(again)).toEqual(bounds);
    expect(getTableColumnWidths(again)).toEqual([inches(1), inches(1), inches(1), inches(3)]);
    expect(getTableRowHeights(again)).toEqual([inches(0.5), inches(0.5), inches(1)]);
    expect(getTableCellParagraphs(getTableCell(again, 0, 0))).toEqual(before);
    expect(getTableCellText(getTableCell(again, 0, 3))).toBe('B');
    expect(getTableCellSpan(getTableCell(again, 0, 3)).rowSpan).toBe(2);
    expect(getTableCellText(getTableCell(again, 2, 0))).toBe('C');
    expect(getTableCellSpan(getTableCell(again, 2, 0)).gridSpan).toBe(3);
    expect(getTableCellText(getTableCell(again, 2, 3))).toBe('D');
    for (let r = 0; r < 2; r++)
      for (let c = 0; c < 3; c++) {
        expect(getTableCellSpan(getTableCell(again, r, c))).toEqual({
          rowSpan: 1,
          gridSpan: 1,
          hMerge: false,
          vMerge: false,
        });
        if (r || c) expect(getTableCellText(getTableCell(again, r, c))).toBe('');
      }
  });

  it('splits an existing merge across new and existing grid boundaries without reviving hidden text', async () => {
    const { p, table } = fixture();
    mergeTableCells(table, { row: 0, col: 0, rowSpan: 2, colSpan: 2 });
    splitTableCell(table, 0, 0, { rows: 3, columns: 3 });
    const again = getSlideTables(
      getSlides(await loadPresentation(await savePresentation(p)))[0]!,
    )[0]!;
    expect(getTableColumnWidths(again)).toEqual([inches(2), inches(1), inches(1), inches(2)]);
    expect(getTableRowHeights(again).reduce((a, b) => a + b, 0)).toBe(inches(2));
    const visible = [];
    for (let r = 0; r < 4; r++)
      for (let c = 0; c < 4; c++) {
        const cell = getTableCell(again, r, c),
          span = getTableCellSpan(cell);
        if (!span.hMerge && !span.vMerge) visible.push(getTableCellText(cell));
      }
    expect(visible).toEqual(['日本語🙂', '', '', '', '', '', '', '', '']);
  });

  it('rejects invalid dimensions and covered cells without changes', async () => {
    const { p, table } = fixture();
    mergeTableCells(table, { row: 0, col: 0, rowSpan: 1, colSpan: 2 });
    const before = await savePresentation(p);
    for (const count of [0, -1, 1.5, NaN, Infinity, 101])
      expect(() => splitTableCell(table, 0, 0, { rows: count, columns: 2 })).toThrow();
    expect(() => splitTableCell(table, 0, 1, { rows: 2, columns: 1 })).toThrow(/anchor/);
    expect(await savePresentation(p)).toEqual(before);
  });

  (isSchemaValidationAvailable() ? it : it.skip)('writes schema-valid split cells', async () => {
    const { p, table } = fixture();
    splitTableCell(table, 0, 0, { rows: 2, columns: 3 });
    expectSchemaValid(
      strFromU8(unzipSync(await savePresentation(p))['ppt/slides/slide1.xml']!),
      'pml',
    );
  });
});
