// `findSlidesByLayoutPartName(pres, layoutPartName)` — finds slides
// keyed on the layout's package part name (stable across locales).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayout,
  findSlidesByLayoutPartName,
  getSlideLayoutPartName,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findSlidesByLayoutPartName', () => {
  it('returns slides whose layout part name matches', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content');
    if (!layout) throw new Error('expected Title and Content layout');
    addSlide(pres, { layout });
    addSlide(pres, { layout });
    const partName = getSlideLayoutPartName(layout);
    const matches = findSlidesByLayoutPartName(pres, partName);
    expect(matches.length).toBe(2);
  });

  it('returns an empty array for an unknown part name', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(findSlidesByLayoutPartName(pres, '/ppt/slideLayouts/slideLayout999.xml')).toEqual([]);
  });
});
