// getSlideMediaPartNames — every media part path the slide references.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideImage,
  getSlideMediaPartNames,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

describe('fn API: getSlideMediaPartNames', () => {
  it('lists every media path the slide references', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideImage(slide, PNG, {
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    addSlideImage(slide, PNG, {
      x: inches(2),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    const names = getSlideMediaPartNames(slide);
    expect(names.length).toBeGreaterThanOrEqual(2);
    for (const name of names) {
      expect(name.startsWith('/ppt/media/')).toBe(true);
    }
  });

  it('returns an empty list on a media-free slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[1]!;
    expect(getSlideMediaPartNames(slide)).toEqual([]);
  });
});
