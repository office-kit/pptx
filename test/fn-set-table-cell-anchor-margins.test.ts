// Writers for `<a:tcPr anchor>` and `<a:tcPr marL/marR/marT/marB>`.
// Pair the existing `getTableCellAnchor` / `getTableCellMargins`
// readers.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTable,
  getSlides,
  getSlideShapes,
  getTableCell,
  getTableCellAnchor,
  getTableCellMargins,
  inches,
  loadPresentation,
  savePresentation,
  setTableCellAnchor,
  setTableCellMargins,
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

describe('fn API: setTableCellAnchor', () => {
  it('round-trips every variant', async () => {
    const { table } = await addDemo();
    const cell = getTableCell(table, 0, 0);
    for (const anchor of ['top', 'center', 'bottom'] as const) {
      setTableCellAnchor(cell, anchor);
      expect(getTableCellAnchor(cell)).toBe(anchor);
    }
  });

  it('survives save/reload', async () => {
    const { pres, table } = await addDemo();
    setTableCellAnchor(getTableCell(table, 1, 1), 'bottom');
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const reShapes = getSlideShapes(getSlides(reloaded)[0]!);
    const tbl = reShapes[reShapes.length - 1]!;
    expect(getTableCellAnchor(getTableCell(tbl, 1, 1))).toBe('bottom');
  });

  it('clears the attribute when set to null', async () => {
    const { table } = await addDemo();
    const cell = getTableCell(table, 0, 0);
    setTableCellAnchor(cell, 'top');
    setTableCellAnchor(cell, null);
    expect(getTableCellAnchor(cell)).toBeNull();
  });
});

describe('fn API: setTableCellMargins', () => {
  it('round-trips per-side margins through save/reload', async () => {
    const { pres, table } = await addDemo();
    const cell = getTableCell(table, 0, 0);
    setTableCellMargins(cell, { left: 50000, right: 60000, top: 70000, bottom: 80000 });
    expect(getTableCellMargins(cell)).toEqual({
      left: 50000,
      right: 60000,
      top: 70000,
      bottom: 80000,
    });

    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const reShapes = getSlideShapes(getSlides(reloaded)[0]!);
    const tbl = reShapes[reShapes.length - 1]!;
    expect(getTableCellMargins(getTableCell(tbl, 0, 0))).toEqual({
      left: 50000,
      right: 60000,
      top: 70000,
      bottom: 80000,
    });
  });

  it('sides set to null/undefined are stripped (fall back to defaults)', async () => {
    const { table } = await addDemo();
    const cell = getTableCell(table, 0, 0);
    setTableCellMargins(cell, { left: 100000, right: 100000, top: 50000, bottom: 50000 });
    setTableCellMargins(cell, { left: 90000 });
    expect(getTableCellMargins(cell)).toEqual({
      left: 90000,
      right: null,
      top: null,
      bottom: null,
    });
  });

  it('null clears every side', async () => {
    const { table } = await addDemo();
    const cell = getTableCell(table, 0, 0);
    setTableCellMargins(cell, { left: 50000, right: 50000, top: 50000, bottom: 50000 });
    setTableCellMargins(cell, null);
    expect(getTableCellMargins(cell)).toEqual({
      left: null,
      right: null,
      top: null,
      bottom: null,
    });
  });
});
