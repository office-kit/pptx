// getAllTables — every table paired with its slide index.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTable,
  getAllTables,
  getSlides,
  inches,
  isTableShape,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getAllTables', () => {
  it('returns an empty list when no slides have tables', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getAllTables(pres)).toEqual([]);
  });

  it('pairs each table with the slide it lives on', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [first, second] = getSlides(pres);
    addSlideTable(first!, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      rows: [['a', 'b']],
    });
    addSlideTable(second!, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      rows: [['c', 'd']],
    });
    addSlideTable(second!, {
      x: inches(4),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      rows: [['e', 'f']],
    });

    const entries = getAllTables(pres);
    expect(entries.length).toBe(3);
    expect(entries.map((e) => e.slideIndex)).toEqual([0, 1, 1]);
    for (const entry of entries) expect(isTableShape(entry.table)).toBe(true);
  });
});
