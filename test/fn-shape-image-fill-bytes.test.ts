// getShapeImageFillBytes — read back the bytes of a shape's image fill.
//
// Counterpart to `getShapeImageBytes` (which covers `<p:pic>` shapes
// only). When `setShapeImageFill` is used on a regular shape, the
// `<a:blipFill>` lives inside `<p:spPr>`; this getter follows that
// `r:embed` rel and returns the underlying media bytes.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  getShapeImageFillBytes,
  getSlides,
  inches,
  loadPresentation,
  setShapeImageFill,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

// 1×1 transparent PNG.
const PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

describe('fn API: getShapeImageFillBytes', () => {
  it('returns the bytes of an image-filled shape', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0), y: inches(0), w: inches(2), h: inches(2),
    });
    setShapeImageFill(shape, PNG);

    const bytes = getShapeImageFillBytes(shape);
    expect(bytes).not.toBeNull();
    expect(bytes!.byteLength).toBe(PNG.byteLength);
    expect(Array.from(bytes!)).toEqual(Array.from(PNG));
  });

  it('returns null when the shape has no image fill', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    expect(getShapeImageFillBytes(shape)).toBeNull();
  });

  it('returns null for a picture shape (use getShapeImageBytes instead)', async () => {
    const { addSlideImage, getShapeImageBytes } = await import('../src/api/index.ts');
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const pic = addSlideImage(slide, PNG, {
      x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    expect(getShapeImageFillBytes(pic)).toBeNull();
    // Sanity: the picture-shape getter still works.
    expect(getShapeImageBytes(pic)).not.toBeNull();
  });
});
