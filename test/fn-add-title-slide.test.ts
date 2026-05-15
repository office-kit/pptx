// addTitleSlide — sugar that adds a slide and sets its title.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addTitleSlide,
  getSlideCount,
  getSlideLayout,
  getSlideLayoutType,
  getSlideTitle,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: addTitleSlide', () => {
  it('appends a slide with the requested title', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const before = getSlideCount(pres);
    const slide = addTitleSlide(pres, 'Quarterly Results');
    expect(getSlideCount(pres)).toBe(before + 1);
    expect(getSlideTitle(slide)).toBe('Quarterly Results');
  });

  it('prefers a title-typed layout', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const slide = addTitleSlide(pres, 'Heading');
    const layout = getSlideLayout(slide)!;
    expect(['title', 'obj']).toContain(getSlideLayoutType(layout));
  });
});
