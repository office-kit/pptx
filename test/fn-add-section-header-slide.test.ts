// addSectionHeaderSlide — sugar for the section-divider pattern.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSectionHeaderSlide,
  getSlideLayout,
  getSlideLayoutType,
  getSlideTitle,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: addSectionHeaderSlide', () => {
  it('adds a slide with the given title', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const slide = addSectionHeaderSlide(pres, 'Part 2: Roadmap');
    expect(getSlideTitle(slide)).toBe('Part 2: Roadmap');
  });

  it('prefers the secHead-typed layout when present', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const slide = addSectionHeaderSlide(pres, 'Divider');
    const layout = getSlideLayout(slide)!;
    expect(['secHead', 'title']).toContain(getSlideLayoutType(layout));
  });
});
