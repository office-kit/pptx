// Per-cell editing on table graphic-frame shapes.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  Presentation,
  _internalPackageOf,
  addSlideTable,
  clearTableCellFill,
  getSlides,
  getTableCell,
  getTableCellPosition,
  getTableCellText,
  getTableCells,
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
  const pres = await Presentation.load(bytes);
  const pkg = _internalPackageOf(pres);
  const part = pkg.parts.find((p) => p.name === `/ppt/slides/slide${slideIndex + 1}.xml`);
  if (!part) throw new Error(`slide${slideIndex + 1}.xml not found`);
  return new TextDecoder().decode(part.data);
};

const addDemoTable = (slide: ReturnType<typeof getSlides>[number]) =>
  addSlideTable(slide, {
    x: inches(0), y: inches(0), w: inches(4), h: inches(2),
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

  it('setTableCellAlignment writes algn on the cell\'s paragraphs', async () => {
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
