// getTableCellFill — read back a cell's solid fill color.
//
// Counterpart to setTableCellFill / clearTableCellFill. Returns
// '#RRGGBB' for sRGB colors, 'scheme:<token>' for scheme colors, or
// null when no solid fill is set (or the cell has no <a:tcPr>).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTable,
  clearTableCellFill,
  getSlides,
  getTableCell,
  getTableCellFill,
  inches,
  loadPresentation,
  setTableCellFill,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const tinyTable = (slide: ReturnType<typeof getSlides>[number]) =>
  addSlideTable(slide, {
    x: inches(0), y: inches(0), w: inches(4), h: inches(2),
    rows: [['', ''], ['', '']],
  });

describe('fn API: getTableCellFill', () => {
  it('returns null on a freshly-built cell', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = tinyTable(slide);
    const cell = getTableCell(table, 0, 0)!;
    expect(getTableCellFill(cell)).toBeNull();
  });

  it('round-trips sRGB color', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = tinyTable(slide);
    const cell = getTableCell(table, 0, 0)!;
    setTableCellFill(cell, '#FFAA00');
    expect(getTableCellFill(cell)).toBe('#FFAA00');

    clearTableCellFill(cell);
    expect(getTableCellFill(cell)).toBeNull();
  });

  it('returns null after clearing a previously-filled cell', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = tinyTable(slide);
    const cell = getTableCell(table, 1, 1)!;
    setTableCellFill(cell, '#112233');
    expect(getTableCellFill(cell)).toBe('#112233');
    clearTableCellFill(cell);
    expect(getTableCellFill(cell)).toBeNull();
  });
});
