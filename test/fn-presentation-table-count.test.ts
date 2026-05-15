// getPresentationTableCount — fast counter for tables across the deck.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTable,
  getAllTables,
  getPresentationTableCount,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getPresentationTableCount', () => {
  it('matches getAllTables length', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [first, second] = getSlides(pres);
    addSlideTable(first!, {
      x: inches(0), y: inches(0), w: inches(4), h: inches(2),
      rows: [['a', 'b']],
    });
    addSlideTable(second!, {
      x: inches(0), y: inches(0), w: inches(4), h: inches(2),
      rows: [['c', 'd']],
    });
    addSlideTable(second!, {
      x: inches(0), y: inches(3), w: inches(4), h: inches(2),
      rows: [['e', 'f']],
    });
    expect(getPresentationTableCount(pres)).toBe(getAllTables(pres).length);
    expect(getPresentationTableCount(pres)).toBeGreaterThanOrEqual(3);
  });

  it('returns 0 on a deck with no tables', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getPresentationTableCount(pres)).toBe(0);
  });
});
