// Per-cell editing on table graphic-frame shapes.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import {
  addSlideTable,
  clearTableCellFill,
  getSlideXmlString,
  getSlides,
  getTableCell,
  getTableCellPosition,
  getTableCellText,
  getTableCellParagraphs,
  setTableCellParagraphs,
  getTableCells,
  getSlideTables,
  inches,
  loadPresentation,
  savePresentation,
  setTableCellAlignment,
  setTableCellFill,
  setTableCellText,
  setTableCellTextFormat,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const slideXml = async (bytes: Uint8Array, slideIndex: number): Promise<string> => {
  const pres = await loadPresentation(bytes);
  return getSlideXmlString(getSlides(pres)[slideIndex]!);
};

const addDemoTable = (slide: ReturnType<typeof getSlides>[number]) =>
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

describe('fn API: table cell access', () => {
  it('getTableCells returns a 2D grid of handles', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addDemoTable(slide);
    const cells = getTableCells(table);
    expect(cells).toHaveLength(2);
    expect(cells[0]).toHaveLength(2);
    expect(getTableCellText(cells[0]![0]!)).toBe('A');
    expect(getTableCellText(cells[1]![1]!)).toBe('D');
    expect(getTableCellPosition(cells[1]![1]!)).toEqual({ row: 1, col: 1 });
  });

  it('setTableCellText replaces one cell only', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addDemoTable(slide);
    const cell = getTableCell(table, 0, 1);
    setTableCellText(cell, 'B'.padStart(5, 'X'));
    expect(getTableCellText(cell)).toBe('XXXXB');
    // Other cells unchanged.
    expect(getTableCellText(getTableCell(table, 0, 0))).toBe('A');
  });

  it('setTableCellFill paints one cell background', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addDemoTable(slide);
    setTableCellFill(getTableCell(table, 0, 0), '#FF0000');
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('FF0000');
  });

  it('setTableCellTextFormat applies to the cell only', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addDemoTable(slide);
    setTableCellTextFormat(getTableCell(table, 1, 0), { bold: true, color: '#00FF00' });
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toMatch(/<a:rPr[^>]*b="1"/);
    expect(xml).toContain('00FF00');
  });

  it("setTableCellAlignment writes algn on the cell's paragraphs", async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addDemoTable(slide);
    setTableCellAlignment(getTableCell(table, 0, 0), 'center');
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toMatch(/<a:pPr[^>]*algn="ctr"/);
  });

  it('clearTableCellFill removes a previously-set fill', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addDemoTable(slide);
    const cell = getTableCell(table, 0, 0);
    setTableCellFill(cell, '#FF0000');
    expect(await slideXml(await savePresentation(pres), 0)).toContain('FF0000');
    clearTableCellFill(cell);
    // No FF0000 from this cell. (Other cells don't carry this color in
    // the demo table, so the substring should disappear entirely.)
    expect(await slideXml(await savePresentation(pres), 0)).not.toContain('FF0000');
  });

  it('throws on out-of-range cells and non-table shapes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addDemoTable(slide);
    expect(() => getTableCell(table, 9, 0)).toThrow(RangeError);
    expect(() => getTableCell(table, 0, 9)).toThrow(RangeError);
  });

  it('getTableDimensions reports row + column counts', async () => {
    const { getTableDimensions } = await import('../src/api/index.ts');
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addDemoTable(slide);
    expect(getTableDimensions(table)).toEqual({ rows: 2, cols: 2 });
  });
});

it('preserves mixed runs and paragraph properties when editing cell text', async () => {
  const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
  const table = addDemoTable(getSlides(pres)[0]!);
  const cell = getTableCell(table, 0, 0);
  setTableCellParagraphs(cell, [
    {
      align: 'center',
      runs: [
        { text: 'Hello ', format: { bold: true } },
        { text: '日本語🌎', format: { italic: true } },
      ],
    },
  ]);
  setTableCellText(cell, 'Hello 日本語🌎!', { preserveFormatting: true });
  const paragraph = getTableCellParagraphs(cell)[0]!;
  expect(paragraph.align).toBe('center');
  expect(paragraph.elements[0]!.format?.bold).toBe(true);
  expect(paragraph.elements[1]!.format?.italic).toBe(true);
  expect(getTableCellText(cell)).toBe('Hello 日本語🌎!');
  const xml = await slideXml(await savePresentation(pres), 0);
  expect(xml).toContain('日本語🌎');
});

it('reads and preserves soft breaks and fields in table cell text', async () => {
  const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
  addDemoTable(getSlides(pres)[0]!);
  const entries = unzipSync(await savePresentation(pres));
  const path = 'ppt/slides/slide1.xml';
  entries[path] = strToU8(
    strFromU8(entries[path]!).replace(
      '<a:t>A</a:t></a:r>',
      '<a:t>A</a:t></a:r><a:br/><a:fld id="{00000000-0000-0000-0000-000000000001}" type="slidenum"><a:t>12</a:t></a:fld>',
    ),
  );
  const loaded = await loadPresentation(zipSync(entries));
  const cell = getTableCell(getSlideTables(getSlides(loaded)[0]!)[0]!, 0, 0);
  expect(getTableCellText(cell)).toBe('A\n12');
  setTableCellText(cell, 'A\n12!', { preserveFormatting: true });
  expect(getTableCellText(cell)).toBe('A\n12!');
  expect(getTableCellParagraphs(cell)[0]!.elements.map((e) => e.kind)).toEqual([
    'r',
    'br',
    'fld',
    'r',
  ]);
});
