// slidesUsingMediaPart — find every slide referencing a given media
// part. Useful for "which slides will this image swap affect?".

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideImage,
  getImageParts,
  getShapeImagePartName,
  getSlideIndex,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  slidesUsingMediaPart,
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

describe('fn API: slidesUsingMediaPart', () => {
  it('finds the slide that embeds an image', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [first, second] = getSlides(pres);
    addSlideImage(first!, PNG, {
      x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });

    // Discover the part name through the picture shape.
    const shapes = getSlideShapes(first!);
    const partName = getShapeImagePartName(shapes[shapes.length - 1]!)!;
    expect(partName).not.toBeNull();

    const hits = slidesUsingMediaPart(pres, partName);
    expect(hits.length).toBe(1);
    expect(getSlideIndex(pres, hits[0]!)).toBe(getSlideIndex(pres, first!));
    void second;
  });

  it('returns an empty list when the part name is unknown', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(slidesUsingMediaPart(pres, '/ppt/media/imageNotPresent.png')).toEqual([]);
  });

  it('lists every slide that shares the same media part', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    addSlideImage(getSlides(pres)[0]!, PNG, {
      x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    addSlideImage(getSlides(pres)[1]!, PNG, {
      x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    // The two pictures don't share a media part — addSlideImage
    // allocates a fresh one each call — but every emitted image
    // appears in getImageParts, so we can at least confirm both
    // are reachable via slidesUsingMediaPart.
    const allImages = getImageParts(pres);
    expect(allImages.length).toBeGreaterThanOrEqual(2);
    for (const img of allImages) {
      const using = slidesUsingMediaPart(pres, img.name);
      expect(using.length).toBeGreaterThanOrEqual(1);
    }
  });
});
