// getOrphanMediaPartNames — media parts no rel references.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideImage,
  compactPackage,
  getOrphanMediaPartNames,
  getShapeImagePartName,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  removeShape,
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

describe('fn API: getOrphanMediaPartNames', () => {
  it('returns an empty list when every media part is referenced', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    addSlideImage(getSlides(pres)[0]!, PNG, {
      x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    expect(getOrphanMediaPartNames(pres)).toEqual([]);
  });

  it('identifies media parts orphaned by removeSlide', async () => {
    const { removeSlide } = await import('../src/api/index.ts');
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const pic = addSlideImage(slide, PNG, {
      x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    const partName = getShapeImagePartName(pic)!;
    expect(partName).not.toBeNull();

    // Drop the whole slide (which removes its .rels) — the media
    // part is now unreachable but still in the package.
    removeSlide(pres, slide);
    expect(getOrphanMediaPartNames(pres)).toContain(partName);

    // compactPackage removes the orphan.
    compactPackage(pres);
    expect(getOrphanMediaPartNames(pres)).toEqual([]);
    void getSlideShapes;
    void removeShape;
  });
});
