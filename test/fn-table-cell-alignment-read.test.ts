// getTableCellAlignment — read back a cell's first-paragraph algn.
// Counterpart to setTableCellAlignment.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTable,
  getSlides,
  getTableCell,
  getTableCellAlignment,
  inches,
  loadPresentation,
  setTableCellAlignment,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getTableCellAlignment', () => {
  it('returns null on a freshly-built cell', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addSlideTable(slide, {
      x: inches(0), y: inches(0), w: inches(4), h: inches(2),
      rows: [['a', 'b']],
    });
    const cell = getTableCell(table, 0, 0)!;
    expect(getTableCellAlignment(cell)).toBeNull();
  });

  it('reflects each ParagraphAlignment value', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addSlideTable(slide, {
      x: inches(0), y: inches(0), w: inches(4), h: inches(2),
      rows: [['a', 'b', 'c', 'd']],
    });

    const c0 = getTableCell(table, 0, 0)!;
    setTableCellAlignment(c0, 'l');
    expect(getTableCellAlignment(c0)).toBe('l');

    const c1 = getTableCell(table, 0, 1)!;
    setTableCellAlignment(c1, 'ctr');
    expect(getTableCellAlignment(c1)).toBe('ctr');

    const c2 = getTableCell(table, 0, 2)!;
    setTableCellAlignment(c2, 'r');
    expect(getTableCellAlignment(c2)).toBe('r');

    const c3 = getTableCell(table, 0, 3)!;
    setTableCellAlignment(c3, 'just');
    expect(getTableCellAlignment(c3)).toBe('just');
  });
});
