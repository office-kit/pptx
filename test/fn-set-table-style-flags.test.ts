// `setTableStyleFlags` — partial-update writer for `<a:tblPr>` boolean
// style flags (firstRow / lastRow / firstCol / lastCol / bandRow /
// bandCol). Pairs the existing `getTableStyleFlags` reader.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTable,
  getSlides,
  getSlideShapes,
  getTableStyleFlags,
  inches,
  loadPresentation,
  savePresentation,
  setTableStyleFlags,
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

describe('fn API: setTableStyleFlags', () => {
  it('round-trips a single flag through save/reload', async () => {
    const { pres, table } = await addDemo();
    setTableStyleFlags(table, { firstRow: true, bandRow: true });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const reShapes = getSlideShapes(getSlides(reloaded)[0]!);
    const tbl = reShapes[reShapes.length - 1]!;
    const flags = getTableStyleFlags(tbl);
    expect(flags.firstRow).toBe(true);
    expect(flags.bandRow).toBe(true);
    expect(flags.lastRow).toBe(false);
    expect(flags.bandCol).toBe(false);
  });

  it('omitted keys are left untouched; false strips the attribute', async () => {
    const { table } = await addDemo();
    setTableStyleFlags(table, { firstRow: true, lastRow: true });
    setTableStyleFlags(table, { firstRow: false });
    const flags = getTableStyleFlags(table);
    expect(flags.firstRow).toBe(false);
    expect(flags.lastRow).toBe(true);
  });

  it('all six flags can be set simultaneously', async () => {
    const { table } = await addDemo();
    setTableStyleFlags(table, {
      firstRow: true,
      lastRow: true,
      firstCol: true,
      lastCol: true,
      bandRow: true,
      bandCol: true,
    });
    expect(getTableStyleFlags(table)).toEqual({
      firstRow: true,
      lastRow: true,
      firstCol: true,
      lastCol: true,
      bandRow: true,
      bandCol: true,
    });
  });

  it('throws on non-table shapes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const nonTable = getSlideShapes(slide)[0]!;
    expect(() => setTableStyleFlags(nonTable, { firstRow: true })).toThrow(
      /shape is not a table graphic frame/,
    );
  });
});
