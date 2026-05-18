// `getUnusedSlideLayouts(pres)` — layouts no slide references.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayout,
  getSlideLayoutName,
  getSlideLayouts,
  getUnusedSlideLayouts,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getUnusedSlideLayouts', () => {
  it('returns an empty array when every layout has at least one slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layouts = getSlideLayouts(pres);
    // Add a slide per layout so none is unused.
    for (const layout of layouts) addSlide(pres, { layout });
    expect(getUnusedSlideLayouts(pres)).toEqual([]);
  });

  it('lists layouts no slide references', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content');
    if (!layout) throw new Error('expected Title and Content layout');
    addSlide(pres, { layout });
    const unused = getUnusedSlideLayouts(pres);
    const unusedNames = unused.map((l) => getSlideLayoutName(l));
    expect(unusedNames).not.toContain('Title and Content');
    expect(unusedNames.length).toBeGreaterThan(0);
  });
});
