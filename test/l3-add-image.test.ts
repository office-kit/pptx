// Level-3: insert a new picture from raw bytes.
//
// Scenario:
//   1. Load `blank.pptx` and add a Blank-layout slide.
//   2. Call `slide.addImage(bytes, { x, y, w, h })` with a small PNG.
//   3. Save → reload → assert the deck has a picture shape with
//      the expected geometry and the media bytes round-trip.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { Presentation, inches } from '../src/api/index.ts';
import { _internalPackageOf } from '../src/api/presentation.ts';
import { partName } from '../src/internal/opc/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

// Tiny 1×1 transparent PNG.
// prettier-ignore
const PNG_1X1 = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

describe('L3: Slide.addImage', () => {
  it('inserts a picture and round-trips through save/reload', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });

    const picture = slide.addImage(PNG_1X1, {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(2),
    });
    expect(picture.kind).toBe('picture');
    expect(picture.position).toEqual({ x: inches(1), y: inches(1) });
    expect(picture.size).toEqual({ w: inches(2), h: inches(2) });

    const reloaded = await Presentation.load(await pres.save());
    const reSlide = reloaded.slides[0];
    expect(reSlide).toBeDefined();
    const rePicture = reSlide?.shapes.find((s) => s.kind === 'picture');
    expect(rePicture).toBeDefined();
    expect(rePicture?.position).toEqual({ x: inches(1), y: inches(1) });
  });

  it('creates a fresh /ppt/media/imageN.png part with the new bytes', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    slide.addImage(PNG_1X1, {
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });

    const reloaded = await Presentation.load(await pres.save());
    const pkg = _internalPackageOf(reloaded);
    const mediaPart = pkg.getPart(partName('/ppt/media/image1.png'));
    expect(mediaPart).not.toBeNull();
    expect(mediaPart?.data).toEqual(PNG_1X1);
  });

  it('multiple addImage calls allocate distinct media parts', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    slide.addImage(PNG_1X1, { x: inches(0), y: inches(0), w: inches(1), h: inches(1) });
    slide.addImage(PNG_1X1, { x: inches(2), y: inches(0), w: inches(1), h: inches(1) });

    const pkg = _internalPackageOf(pres);
    const mediaParts = pkg.parts.filter((p) => p.name.startsWith('/ppt/media/image'));
    expect(mediaParts.length).toBe(2);
  });

  it('throws when format cannot be detected', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    expect(() =>
      slide.addImage(new Uint8Array([0, 0, 0]), {
        x: inches(0),
        y: inches(0),
        w: inches(1),
        h: inches(1),
      }),
    ).toThrow(/format/);
  });
});
