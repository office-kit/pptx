// Table span / border readers — coverage of the merge attributes
// (`gridSpan`, `rowSpan`, `hMerge`, `vMerge`) and per-cell borders
// (`<a:tcPr><a:ln{L,R,T,B,TlToBr,BlToTr}>`). These complete the table
// read-back surface so renderers can place merged regions and draw the
// borders PowerPoint actually emits.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTable,
  getTableCell,
  getTableCellBorders,
  getTableCellSpan,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getTableCellSpan / getTableCellBorders', () => {
  it('defaults to gridSpan=1 / rowSpan=1 / no merge for unauthored cells', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tbl = addSlideTable(slide, {
      rows: [['a', 'b'], ['c', 'd']],
      x: inches(0), y: inches(0), w: inches(4), h: inches(2),
    });
    const cell = getTableCell(tbl, 0, 0);
    const span = getTableCellSpan(cell);
    expect(span.gridSpan).toBe(1);
    expect(span.rowSpan).toBe(1);
    expect(span.hMerge).toBe(false);
    expect(span.vMerge).toBe(false);
  });

  it('returns an all-null TableCellBorders for cells without authored borders', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tbl = addSlideTable(slide, {
      rows: [['a']],
      x: inches(0), y: inches(0), w: inches(4), h: inches(2),
    });
    const cell = getTableCell(tbl, 0, 0);
    const borders = getTableCellBorders(pres, cell);
    expect(borders.left).toBeNull();
    expect(borders.right).toBeNull();
    expect(borders.top).toBeNull();
    expect(borders.bottom).toBeNull();
    expect(borders.tlToBr).toBeNull();
    expect(borders.blToTr).toBeNull();
  });
});
