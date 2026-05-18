// `setTableStyleId` — writer for `<a:tblPr><a:tableStyleId>`. Pairs the
// existing `getTableStyleId` reader.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTable,
  getSlides,
  getSlideShapes,
  getTableStyleId,
  inches,
  loadPresentation,
  savePresentation,
  setTableStyleId,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const STYLE_GUID = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';

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

describe('fn API: setTableStyleId', () => {
  it('round-trips a style GUID through save/reload', async () => {
    const { pres, table } = await addDemo();
    setTableStyleId(table, STYLE_GUID);
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const reShapes = getSlideShapes(getSlides(reloaded)[0]!);
    const tbl = reShapes[reShapes.length - 1]!;
    expect(getTableStyleId(tbl)).toBe(STYLE_GUID);
  });

  it('null removes the element', async () => {
    const { table } = await addDemo();
    setTableStyleId(table, STYLE_GUID);
    expect(getTableStyleId(table)).toBe(STYLE_GUID);
    setTableStyleId(table, null);
    expect(getTableStyleId(table)).toBeNull();
  });

  it('replaces a prior id on subsequent calls', async () => {
    const { table } = await addDemo();
    setTableStyleId(table, STYLE_GUID);
    setTableStyleId(table, '{00000000-0000-0000-0000-000000000000}');
    expect(getTableStyleId(table)).toBe('{00000000-0000-0000-0000-000000000000}');
  });

  it('throws on non-table shapes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const nonTable = getSlideShapes(slide)[0]!;
    expect(() => setTableStyleId(nonTable, STYLE_GUID)).toThrow(
      /shape is not a table graphic frame/,
    );
  });
});
