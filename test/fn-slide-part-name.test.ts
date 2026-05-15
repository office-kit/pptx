// getSlidePartName / findSlideByPartName — round-trip a slide via
// its OPC part path.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  findSlideByPartName,
  getSlidePartName,
  getSlides,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlidePartName / findSlideByPartName', () => {
  it('returns the /ppt/slides/slideN.xml path for each slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    for (const slide of getSlides(pres)) {
      const path = getSlidePartName(slide);
      expect(path.startsWith('/ppt/slides/slide')).toBe(true);
      expect(path.endsWith('.xml')).toBe(true);
    }
  });

  it('findSlideByPartName round-trips identity', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slides = getSlides(pres);
    for (const slide of slides) {
      const path = getSlidePartName(slide);
      const looked = findSlideByPartName(pres, path);
      expect(looked).toBe(slide);
    }
  });

  it('returns null for a non-existent part name', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(findSlideByPartName(pres, '/ppt/slides/no-such-slide.xml')).toBeNull();
  });
});
