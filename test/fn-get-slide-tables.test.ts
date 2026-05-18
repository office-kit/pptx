// `getSlideTables(slide)` — slide-scoped table-shape listing.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTable,
  addSlideTextBox,
  getSlideTables,
  getSlides,
  inches,
  isTableShape,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideTables', () => {
  it('returns every table on the slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideTable(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      rows: [['A', 'B']],
    });
    addSlideTable(slide, {
      x: inches(0),
      y: inches(2),
      w: inches(4),
      h: inches(2),
      rows: [['C', 'D']],
    });
    // Plus a non-table shape that shouldn't show up.
    addSlideTextBox(slide, {
      x: inches(0),
      y: inches(4),
      w: inches(4),
      h: inches(1),
      text: 'ignored',
    });
    const tables = getSlideTables(slide);
    expect(tables.length).toBe(2);
    expect(tables.every((t) => isTableShape(t))).toBe(true);
  });

  it('returns an empty array on a slide without tables', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    expect(getSlideTables(slide)).toEqual([]);
  });
});
