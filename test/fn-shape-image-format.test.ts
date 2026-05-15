// getShapeImageFormat — detect the image format token from a shape's
// embedded bytes. Works for both picture shapes (<p:pic>) and shapes
// with image fill (<a:blipFill> nested in <p:spPr>).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideImage,
  addSlideShape,
  getShapeImageFormat,
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

// Tiny JPEG (SOI + APP0 + EOI is enough for the magic check).
const JPEG = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00, 0x01, 0xff, 0xd9]);

describe('fn API: getShapeImageFormat', () => {
  it('reports png for a PNG picture', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const pic = addSlideImage(slide, PNG, {
      x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    expect(getShapeImageFormat(pic)).toBe('png');
  });

  it('reports jpeg for a JPEG picture', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const pic = addSlideImage(slide, JPEG, {
      x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    expect(getShapeImageFormat(pic)).toBe('jpeg');
  });

  it('reports png for an image-filled shape', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const rect = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    setShapeImageFill(rect, PNG);
    expect(getShapeImageFormat(rect)).toBe('png');
  });

  it('returns null for a shape with no image', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const rect = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    expect(getShapeImageFormat(rect)).toBeNull();
  });
});
