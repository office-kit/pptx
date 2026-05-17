// Level-3: insert a new picture from raw bytes.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideImage,
  findSlideLayout,
  getMediaParts,
  getShapeKind,
  getShapePosition,
  getShapeSize,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const PNG_1X1 = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

describe('L3: addSlideImage', () => {
  it('inserts a picture and round-trips through save/reload', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = addSlide(pres, { layout });

    const picture = addSlideImage(slide, PNG_1X1, {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(2),
    });
    expect(getShapeKind(picture)).toBe('picture');
    expect(getShapePosition(picture)).toEqual({ x: inches(1), y: inches(1) });
    expect(getShapeSize(picture)).toEqual({ w: inches(2), h: inches(2) });

    const reloaded = await loadPresentation(await savePresentation(pres));
    const reSlide = getSlides(reloaded)[0]!;
    const rePicture = getSlideShapes(reSlide).find((s) => getShapeKind(s) === 'picture');
    expect(rePicture).toBeDefined();
    expect(rePicture && getShapePosition(rePicture)).toEqual({ x: inches(1), y: inches(1) });
  });

  it('creates a fresh /ppt/media/imageN.png part with the new bytes', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = addSlide(pres, { layout });
    addSlideImage(slide, PNG_1X1, {
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });

    const reloaded = await loadPresentation(await savePresentation(pres));
    const media = getMediaParts(reloaded).find((p) => p.name === '/ppt/media/image1.png');
    expect(media).toBeDefined();
    expect(media?.data).toEqual(PNG_1X1);
  });

  it('multiple addSlideImage calls allocate distinct media parts', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = addSlide(pres, { layout });
    addSlideImage(slide, PNG_1X1, { x: inches(0), y: inches(0), w: inches(1), h: inches(1) });
    addSlideImage(slide, PNG_1X1, { x: inches(2), y: inches(0), w: inches(1), h: inches(1) });

    const mediaParts = getMediaParts(pres).filter((p) => p.name.startsWith('/ppt/media/image'));
    expect(mediaParts.length).toBe(2);
  });

  it('throws when format cannot be detected', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = addSlide(pres, { layout });
    expect(() =>
      addSlideImage(slide, new Uint8Array([0, 0, 0]), {
        x: inches(0),
        y: inches(0),
        w: inches(1),
        h: inches(1),
      }),
    ).toThrow(/format/);
  });
});
