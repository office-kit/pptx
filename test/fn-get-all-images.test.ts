// getAllImages — every image-bearing shape paired with its slide index.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideImage,
  addSlideShape,
  getAllImages,
  getSlides,
  hasShapeImage,
  inches,
  loadPresentation,
  setShapeImageFill,
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

describe('fn API: getAllImages', () => {
  it('returns an empty list when nothing has images', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getAllImages(pres)).toEqual([]);
  });

  it('lists pictures and image-filled shapes together', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [first, second] = getSlides(pres);
    addSlideImage(first!, PNG, {
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    const filled = addSlideShape(second!, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    setShapeImageFill(filled, PNG);

    const entries = getAllImages(pres);
    expect(entries.length).toBe(2);
    expect(entries[0]!.slideIndex).toBe(0);
    expect(entries[1]!.slideIndex).toBe(1);
    for (const e of entries) expect(hasShapeImage(e.shape)).toBe(true);
  });
});
