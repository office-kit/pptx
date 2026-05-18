// `setTableCellBorders` — partial-update writer for all 6 sides
// (left/right/top/bottom + tlToBr/blToTr diagonals). Pairs the existing
// `getTableCellBorders` reader.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTable,
  getSlides,
  getSlideShapes,
  getTableCell,
  getTableCellBorders,
  inches,
  loadPresentation,
  savePresentation,
  setTableCellBorders,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const addDemo = async () => {
  const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
  const slide = getSlides(pres)[0]!;
  const table = addSlideTable(slide, {
    x: inches(0),
    y: inches(0),
    w: inches(4),
    h: inches(2),
    rows: [
      ['A', 'B'],
      ['C', 'D'],
    ],
  });
  return { pres, table };
};

describe('fn API: setTableCellBorders', () => {
  it('writes one side and reads it back through save/reload', async () => {
    const { pres, table } = await addDemo();
    const cell = getTableCell(table, 0, 0);
    setTableCellBorders(cell, {
      bottom: { color: '#FF0000', widthEmu: 19050, dash: 'dash' },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const reShapes = getSlideShapes(getSlides(reloaded)[0]!);
    const tbl = reShapes[reShapes.length - 1]!;
    const borders = getTableCellBorders(reloaded, getTableCell(tbl, 0, 0));
    expect(borders.bottom).toEqual({ color: '#FF0000', widthEmu: 19050, dash: 'dash' });
    expect(borders.left).toBeNull();
    expect(borders.right).toBeNull();
    expect(borders.top).toBeNull();
  });

  it('omitted sides are left untouched; null clears just that side', async () => {
    const { pres, table } = await addDemo();
    const cell = getTableCell(table, 0, 0);
    setTableCellBorders(cell, {
      left: { color: '#112233', widthEmu: 12700, dash: null },
      right: { color: '#445566', widthEmu: 12700, dash: null },
    });
    // Partial update: clear `left` only, leave `right` intact.
    setTableCellBorders(cell, { left: null });
    const borders = getTableCellBorders(pres, cell);
    expect(borders.left).toBeNull();
    expect(borders.right).toEqual({ color: '#445566', widthEmu: 12700, dash: null });
  });

  it('null as the sides arg clears every side', async () => {
    const { pres, table } = await addDemo();
    const cell = getTableCell(table, 0, 0);
    setTableCellBorders(cell, {
      left: { color: '#000000', widthEmu: 6350, dash: null },
      right: { color: '#000000', widthEmu: 6350, dash: null },
      top: { color: '#000000', widthEmu: 6350, dash: null },
      bottom: { color: '#000000', widthEmu: 6350, dash: null },
    });
    setTableCellBorders(cell, null);
    const borders = getTableCellBorders(pres, cell);
    expect(borders.left).toBeNull();
    expect(borders.right).toBeNull();
    expect(borders.top).toBeNull();
    expect(borders.bottom).toBeNull();
  });

  it('round-trips diagonal borders', async () => {
    const { pres, table } = await addDemo();
    const cell = getTableCell(table, 0, 0);
    setTableCellBorders(cell, {
      tlToBr: { color: '#00AA00', widthEmu: 19050, dash: null },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const reShapes = getSlideShapes(getSlides(reloaded)[0]!);
    const tbl = reShapes[reShapes.length - 1]!;
    expect(getTableCellBorders(reloaded, getTableCell(tbl, 0, 0)).tlToBr).toEqual({
      color: '#00AA00',
      widthEmu: 19050,
      dash: null,
    });
  });
});
