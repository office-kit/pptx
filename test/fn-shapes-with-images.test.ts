// hasShapeImage — predicate that flags shapes carrying images
// (picture or image-fill).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideImage,
  addSlideShape,
  addSlideTextBox,
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

describe('fn API: hasShapeImage', () => {
  it('predicate is true for picture shapes and image-filled shapes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const pic = addSlideImage(slide, PNG, {
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    expect(hasShapeImage(pic)).toBe(true);

    const filled = addSlideShape(slide, {
      preset: 'rect',
      x: inches(2),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    setShapeImageFill(filled, PNG);
    expect(hasShapeImage(filled)).toBe(true);
  });

  it('predicate is false for plain shapes and textboxes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const rect = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(2),
      w: inches(2),
      h: inches(1),
      text: 'plain',
    });
    expect(hasShapeImage(rect)).toBe(false);
    expect(hasShapeImage(tb)).toBe(false);
  });
});
