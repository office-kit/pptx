// getImageParts — filtered view of getMediaParts for image-only.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideImage,
  getImageParts,
  getMediaParts,
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

describe('fn API: getImageParts', () => {
  it('returns image-typed parts only', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideImage(slide, PNG, {
      x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    addSlideImage(slide, PNG, {
      x: inches(2), y: inches(0), w: inches(1), h: inches(1),
    });
    const images = getImageParts(pres);
    expect(images.length).toBeGreaterThanOrEqual(2);
    for (const part of images) {
      expect(part.contentType.startsWith('image/')).toBe(true);
    }
  });

  it('is a strict subset of getMediaParts', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    addSlideImage(getSlides(pres)[0]!, PNG, {
      x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    const all = getMediaParts(pres);
    const images = getImageParts(pres);
    const allNames = new Set(all.map((p) => p.name));
    for (const img of images) {
      expect(allNames.has(img.name)).toBe(true);
    }
  });

  it('returns empty when no embedded images', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    expect(getImageParts(pres)).toEqual([]);
  });
});
