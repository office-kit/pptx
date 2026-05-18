// `setTableCellTextDirection` — vertical-text writer for table cells.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTable,
  getSlides,
  getSlideShapes,
  getTableCell,
  getTableCellTextDirection,
  inches,
  loadPresentation,
  savePresentation,
  setTableCellTextDirection,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const DIRECTIONS = [
  'vert',
  'vert270',
  'wordArtVert',
  'eaVert',
  'mongolianVert',
  'wordArtVertRtl',
] as const;

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

describe('fn API: setTableCellTextDirection', () => {
  it('round-trips every variant', async () => {
    const { table } = await addDemo();
    const cell = getTableCell(table, 0, 0);
    for (const d of DIRECTIONS) {
      setTableCellTextDirection(cell, d);
      expect(getTableCellTextDirection(cell)).toBe(d);
    }
  });

  it('survives save/reload', async () => {
    const { pres, table } = await addDemo();
    setTableCellTextDirection(getTableCell(table, 0, 1), 'eaVert');
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const reShapes = getSlideShapes(getSlides(reloaded)[0]!);
    const tbl = reShapes[reShapes.length - 1]!;
    expect(getTableCellTextDirection(getTableCell(tbl, 0, 1))).toBe('eaVert');
  });

  it('clears the attribute when set to null or horz', async () => {
    const { table } = await addDemo();
    const cell = getTableCell(table, 0, 0);
    setTableCellTextDirection(cell, 'vert');
    expect(getTableCellTextDirection(cell)).toBe('vert');
    setTableCellTextDirection(cell, null);
    expect(getTableCellTextDirection(cell)).toBeNull();
    setTableCellTextDirection(cell, 'vert');
    setTableCellTextDirection(cell, 'horz');
    expect(getTableCellTextDirection(cell)).toBeNull();
  });
});
