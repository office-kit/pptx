import { describe, expect, it } from 'vitest';
import * as api from '../src/api/index.ts';
import {
  Presentation,
  Slide,
  Table,
  compile,
  type TableCellInfo,
  type TableProps,
} from '../packages/dsl/src/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const BOX = { x: 0.6, y: 1, width: 8, height: 2 } as const;

const compileTable = async (props: Omit<TableProps, keyof typeof BOX>) => {
  const pres = await compile(
    Presentation({ children: Slide({ children: Table({ ...BOX, ...props }) }) }),
  );
  const table = api.getSlideShapes(api.getSlides(pres)[0]!)[0]!;
  return { pres, table, cell: (row: number, col: number) => api.getTableCell(table, row, col) };
};

const runFormats = (cell: api.TableCellData) =>
  api
    .getTableCellParagraphs(cell)
    .flatMap((paragraph) => paragraph.elements)
    .map((run) => run.kind === 'r' && run.format);

const setSides = (pres: api.PresentationData, cell: api.TableCellData) =>
  Object.entries(api.getTableCellBorders(pres, cell))
    .filter(([, border]) => border !== null)
    .map(([side]) => side);

describe('Table rich cells', () => {
  it('keeps string cells working next to cell objects', async () => {
    const { cell } = await compileTable({
      rows: [['Plan', { text: 'Count' }, { paragraphs: [{ runs: [{ text: 'Note' }] }] }]],
    });
    expect([0, 1, 2].map((c) => api.getTableCellText(cell(0, c)))).toEqual([
      'Plan',
      'Count',
      'Note',
    ]);
  });

  // Replacing the runs after the table-wide format was applied lost size and
  // color, so the merged format is the base every run is built on.
  it('formats rich runs on top of cellStyle, headerStyle and styleCell', async () => {
    const rich = (text: string) => ({
      paragraphs: [{ runs: [{ text }, { text: '!', format: { bold: true, color: '#C00000' } }] }],
    });
    const { cell } = await compileTable({
      rows: [
        ['Metric', rich('Head')],
        ['Tickets', rich('372')],
      ],
      cellStyle: { format: { size: 16, color: '#15171C' } },
      headerStyle: { format: { color: '#FFFFFF' } },
      styleCell: ({ row, column }) =>
        row === 1 && column === 1 ? { format: { italic: true } } : undefined,
    });
    expect(runFormats(cell(0, 1))).toEqual([
      { size: 16, color: '#FFFFFF' },
      { size: 16, color: '#C00000', bold: true },
    ]);
    expect(runFormats(cell(1, 1))).toEqual([
      { size: 16, color: '#15171C', italic: true },
      { size: 16, color: '#C00000', italic: true, bold: true },
    ]);
    expect(runFormats(cell(1, 0))).toEqual([{ size: 16, color: '#15171C' }]);
  });

  it('keeps the deck body-text color on a rich cell with no format at all', async () => {
    const { cell } = await compileTable({
      rows: [['plain', { paragraphs: [{ runs: [{ text: 'rich' }] }] }]],
    });
    expect(runFormats(cell(0, 1))).toEqual(runFormats(cell(0, 0)));
    expect(runFormats(cell(0, 0))[0]).toHaveProperty('color');
  });

  it("aligns a rich cell from the cell style; a paragraph's own align wins", async () => {
    const { cell } = await compileTable({
      rows: [
        [{ paragraphs: [{ runs: [{ text: '18' }] }, { align: 'l', runs: [{ text: 'low' }] }] }],
      ],
      styleCell: () => ({ align: 'r', fill: '#FEE08B', anchor: 'center' }),
    });
    expect(api.getTableCellParagraphs(cell(0, 0)).map((p) => p.align)).toEqual(['right', 'left']);
    expect(api.getTableCellFill(cell(0, 0))).toBe('#FEE08B');
    expect(api.getTableCellAnchor(cell(0, 0))).toBe('center');
  });

  it('passes styleCell the joined run text of a rich cell', async () => {
    const seen: TableCellInfo[] = [];
    await compileTable({
      rows: [
        [
          'Plan',
          {
            paragraphs: [
              { runs: [{ text: '372 ' }, { text: '(+26%)' }] },
              { runs: [{ text: 'vs. Q1' }] },
            ],
          },
        ],
      ],
      styleCell: (info) => {
        seen.push(info);
        return undefined;
      },
    });
    expect(seen).toEqual([
      { row: 0, column: 0, value: 'Plan' },
      { row: 0, column: 1, value: '372 (+26%)\nvs. Q1' },
    ]);
  });
});

describe('Table merges', () => {
  it('merges from the top-left cell and leaves covered positions empty', async () => {
    const { cell } = await compileTable({
      rows: [
        [{ text: 'Plan', rowSpan: 2 }, { text: 'Effect', colSpan: 2 }, ''],
        ['', 'Count', 'Note'],
        ['FAQ', '372', 'ok'],
      ],
    });
    expect(api.getTableCellSpan(cell(0, 0))).toMatchObject({ gridSpan: 1, rowSpan: 2 });
    expect(api.getTableCellSpan(cell(0, 1))).toMatchObject({ gridSpan: 2, rowSpan: 1 });
    expect(api.getTableCellSpan(cell(0, 2))).toMatchObject({ hMerge: true });
    expect(api.getTableCellSpan(cell(1, 0))).toMatchObject({ vMerge: true });
    expect(api.getTableCellText(cell(1, 1))).toBe('Count');
  });

  it('merges a block from one cell that carries both spans', async () => {
    const { pres } = await compileTable({
      rows: [
        [{ text: 'Block', rowSpan: 2, colSpan: 2 }, '', 'a'],
        ['', '', 'b'],
        ['c', 'd', 'e'],
      ],
    });
    const reloaded = await api.loadPresentation(await api.savePresentation(pres));
    const table = api.getSlideShapes(api.getSlides(reloaded)[0]!)[0]!;
    const span = (row: number, col: number) =>
      api.getTableCellSpan(api.getTableCell(table, row, col));
    expect(span(0, 0)).toMatchObject({ gridSpan: 2, rowSpan: 2 });
    expect(span(0, 1)).toMatchObject({ hMerge: true });
    expect(span(1, 0)).toMatchObject({ vMerge: true });
    expect(span(1, 1)).toMatchObject({ hMerge: true, vMerge: true });
    expect(span(2, 0)).toMatchObject({ gridSpan: 1, rowSpan: 1, hMerge: false, vMerge: false });
    expect(api.getTableCellText(api.getTableCell(table, 2, 2))).toBe('e');
    if (isSchemaValidationAvailable())
      expectSchemaValid(
        new TextDecoder().decode(api.readPackagePart(reloaded, '/ppt/slides/slide1.xml')!),
        'pml',
      );
  });

  // PowerPoint paints a merged block from its top-left cell alone, fill and all
  // four borders, so a covered position is left bare and styleCell never sees it.
  it('styles the top-left cell of a merge and leaves covered positions bare', async () => {
    const seen: string[] = [];
    const { pres, cell } = await compileTable({
      rows: [
        ['Head', 'A', 'B'],
        [{ text: 'Block', rowSpan: 2, colSpan: 2 }, '', 'x'],
        ['', '', 'y'],
      ],
      cellStyle: { fill: '#F0F0F0', borders: { bottom: { color: '#D5D9E0', width: 0.75 } } },
      stripeFill: '#F3F4F7',
      styleCell: ({ row, column }) => {
        seen.push(`${row}:${column}`);
        return row === 1 && column === 0 ? { borders: { right: { color: '#15171C' } } } : undefined;
      },
    });
    expect(seen).toEqual(['0:0', '0:1', '0:2', '1:0', '1:2', '2:2']);
    expect(api.getTableCellFill(cell(1, 0))).toBe('#F0F0F0');
    expect(setSides(pres, cell(1, 0))).toEqual(['right', 'bottom']);
    for (const [row, col] of [
      [1, 1],
      [2, 0],
      [2, 1],
    ] as const) {
      expect(api.getTableCellFill(cell(row, col))).toBeNull();
      expect(setSides(pres, cell(row, col))).toEqual([]);
    }
    expect(api.getTableCellFill(cell(2, 2))).toBe('#F3F4F7');
    expect(setSides(pres, cell(2, 2))).toEqual(['bottom']);
  });

  it('rejects content in a position a merge covers', async () => {
    await expect(
      compileTable({
        rows: [
          [{ text: 'Effect', colSpan: 2 }, 'lost'],
          ['a', 'b'],
        ],
      }),
    ).rejects.toThrow("Table cell (0, 1) is covered by the merge at (0, 0); write it as ''.");
    await expect(
      compileTable({
        rows: [
          [{ text: 'A', rowSpan: 2 }, 'x'],
          [{ text: 'B' }, 'y'],
        ],
      }),
    ).rejects.toThrow('Table cell (1, 0) is covered by the merge at (0, 0)');
  });

  it('rejects text together with paragraphs, and a cell object with neither', async () => {
    await expect(
      compileTable({
        // @ts-expect-error The type allows one of the two; untyped callers get the runtime error.
        rows: [[{ text: 'a', paragraphs: [{ runs: [{ text: 'b' }] }] }]],
      }),
    ).rejects.toThrow('Table cell (0, 0) accepts either text or paragraphs.');
    await expect(
      compileTable({
        // @ts-expect-error An empty cell is the string ''; untyped callers get the runtime error.
        rows: [['a', {}]],
      }),
    ).rejects.toThrow('Table cell (0, 1) accepts either text or paragraphs.');
  });

  it('passes span errors through from the core', async () => {
    await expect(compileTable({ rows: [[{ text: 'a', colSpan: 3 }, '']] })).rejects.toThrow(
      /exceed/,
    );
    await expect(compileTable({ rows: [[{ text: 'a', colSpan: 1.5 }, '']] })).rejects.toThrow(
      /integers/,
    );
  });

  (isSchemaValidationAvailable() ? it : it.skip)('emits schema-valid slide XML', async () => {
    const { pres } = await compileTable({
      rows: [
        [{ text: 'Plan', rowSpan: 2 }, { text: 'Effect', colSpan: 2 }, ''],
        ['', 'Count', { paragraphs: [{ runs: [{ text: 'Note' }] }] }],
        ['FAQ', '372', 'ok'],
      ],
      cellStyle: { borders: { bottom: { color: '#D5D9E0', width: 0.75 } } },
      headerStyle: { fill: '#15171C', format: { color: '#FFFFFF' } },
      styleCell: ({ value }) => (value === '372' ? { fill: '#D73027', align: 'r' } : undefined),
    });
    const reloaded = await api.loadPresentation(await api.savePresentation(pres));
    expect(api.validatePresentation(reloaded).filter((i) => i.severity === 'error')).toEqual([]);
    expectSchemaValid(
      new TextDecoder().decode(api.readPackagePart(reloaded, '/ppt/slides/slide1.xml')!),
      'pml',
    );
  });
});
