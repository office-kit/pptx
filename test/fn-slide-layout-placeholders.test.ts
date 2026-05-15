// getSlideLayoutPlaceholders — enumerate placeholder slots on a layout.
//
// The blank.pptx fixture ships PowerPoint's default eleven layouts;
// "Title and Content" has a title placeholder + a body placeholder,
// "Blank" has none. The helper filters out non-placeholder shapes.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  findSlideLayout,
  getSlideLayoutPlaceholders,
  getSlideLayouts,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideLayoutPlaceholders', () => {
  it('returns a title-typed entry on "Title and Content"', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    const phs = getSlideLayoutPlaceholders(layout);
    // Must have at least one placeholder, and at least one of them is
    // the title slot (PowerPoint may emit `type="title"` or `ctrTitle`).
    expect(phs.length).toBeGreaterThan(0);
    const hasTitle = phs.some((p) => p.type === 'title' || p.type === 'ctrTitle');
    expect(hasTitle).toBe(true);
  });

  it('the "Blank" layout has only date/footer/slide-number placeholders', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const phs = getSlideLayoutPlaceholders(blank);
    // PowerPoint's default Blank layout ships three placeholders for
    // date / slide-number / footer that the user can opt into via
    // "Insert > Header & Footer", but no title or body slot.
    const types = phs.map((p) => p.type).sort();
    expect(types).toEqual(['dt', 'ftr', 'sldNum']);
    expect(phs.some((p) => p.type === 'title')).toBe(false);
    expect(phs.some((p) => p.type === 'body')).toBe(false);
  });

  it('walks every layout without throwing', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layouts = getSlideLayouts(pres);
    expect(layouts.length).toBeGreaterThan(0);
    for (const layout of layouts) {
      const phs = getSlideLayoutPlaceholders(layout);
      for (const p of phs) {
        // Either the type is set or the idx is set (or both); never
        // both null (the implementation guarantees this).
        expect(p.type !== null || p.idx !== null).toBe(true);
        expect(typeof p.name).toBe('string');
      }
    }
  });
});
