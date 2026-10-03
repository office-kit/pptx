import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import {
  addSlideTable,
  getSlides,
  getTableCells,
  getTableCellText,
  inches,
  loadPresentation,
  mergeTableCells,
  setTableCellTextFormat,
  getTableCellParagraphs,
} from '../src/api/index.ts';
import {
  parseTableClipboard,
  serializeTableClipboard,
  copyTableCellValues,
  canPasteTableCells,
  pasteTableCells,
} from '../site/src/lib/editor/core/table-clipboard.ts';

async function table() {
  const pres = await loadPresentation(
    await readFile(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url)),
  );
  return addSlideTable(getSlides(pres)[0]!, {
    x: inches(0),
    y: inches(0),
    w: inches(4),
    h: inches(2),
    rows: [
      ['A', 'B'],
      ['C', 'D'],
    ],
  });
}

describe('spreadsheet clipboard', () => {
  it('parses quoted tabs, multiline Japanese, escaped quotes and terminal CRLF', () => {
    expect(parseTableClipboard('"日本語\nEnglish"\t"a\tb"\r\n"Say ""hello"""\t\r\n')).toEqual([
      ['日本語\nEnglish', 'a\tb'],
      ['Say "hello"', ''],
    ]);
    expect(parseTableClipboard('A\nB\n')).toEqual([['A'], ['B']]);
    expect(parseTableClipboard('A\tB\t')).toEqual([['A', 'B', '']]);
    expect(parseTableClipboard('A\n\n')).toEqual([['A'], ['']]);
    expect(parseTableClipboard('"broken')).toBeNull();
    expect(parseTableClipboard('"closed"extra')).toBeNull();
  });
  it('round-trips copied strings without confusing field delimiters', () => {
    const values = [
      ['日本語\nEnglish', 'a\tb', 'say "Hi"'],
      ['', '', ''],
    ];
    expect(parseTableClipboard(serializeTableClipboard(values))).toEqual(values);
  });
  it('copies complete merged anchors and blanks their covered cells', async () => {
    const target = await table();
    mergeTableCells(target, { row: 0, col: 0, rowSpan: 1, colSpan: 2 });
    expect(
      copyTableCellValues(target, { kind: 'cell', slideIndex: 0, shapeId: 1, row: 0, col: 0 }),
    ).toEqual([['A', '']]);
  });
  it('expands rows and columns and preserves cells outside a ragged paste', async () => {
    const target = await table();
    setTableCellTextFormat(getTableCells(target)[0]![0]!, { bold: true });
    setTableCellTextFormat(getTableCells(target)[1]![1]!, { italic: true });
    const values = [['日本語', 'English'], ['last']];
    expect(canPasteTableCells(target, 1, 1, values)).toBe(true);
    pasteTableCells(target, 1, 1, values);
    expect(getTableCells(target).map((row) => row.map(getTableCellText))).toEqual([
      ['A', 'B', ''],
      ['C', '日本語', 'English'],
      ['', 'last', ''],
    ]);
    expect(
      getTableCellParagraphs(getTableCells(target)[0]![0]!)[0]!.elements[0]!.format?.bold,
    ).toBe(true);
    expect(
      getTableCellParagraphs(getTableCells(target)[1]![1]!)[0]!.elements[0]!.format?.italic,
    ).toBe(true);
  });
  it('rejects intersecting or expanding merged tables before mutation', async () => {
    const target = await table();
    mergeTableCells(target, { row: 0, col: 0, rowSpan: 1, colSpan: 2 });
    expect(canPasteTableCells(target, 0, 0, [['X', 'Y']])).toBe(false);
    expect(canPasteTableCells(target, 1, 0, [['X', 'Y']])).toBe(true);
    expect(canPasteTableCells(target, 1, 0, [['X', 'Y'], ['Z']])).toBe(false);
    expect(getTableCellText(getTableCells(target)[1]![0]!)).toBe('C');
  });
});
