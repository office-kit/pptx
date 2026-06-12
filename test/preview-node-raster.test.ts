// Smoke tests for the Node rasterization path (`pptx-kit-preview/node`).
//
// The pipeline: renderSlideToSvg (svg text mode) → resvg → PNG / RGBA.
// We verify dimensions, the PNG magic bytes, that the output is not entirely
// blank white, and that two renders of the same slide are byte-identical.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideShape,
  addSlideTextBox,
  findSlideLayout,
  getSlideSize,
  inches,
  loadPresentation,
  setShapeFill,
} from '../src/api/index.ts';
import { renderSlideToImage, renderSlideToRgba } from '../packages/preview/src/node.ts';

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));

const buildTestSlide = async () => {
  const pres = await loadPresentation(await readFile(fixturePath));
  const layout = findSlideLayout(pres, 'Blank');
  if (!layout) throw new Error('Blank layout not found');
  const slide = addSlide(pres, { layout });
  // Colored background rect so the rendered image is not entirely white.
  const bg = addSlideShape(slide, {
    preset: 'rect',
    x: inches(0),
    y: inches(0),
    w: inches(10),
    h: inches(7.5),
  });
  setShapeFill(bg, '#3B82F6');
  // A text box for additional content variety.
  addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(8),
    h: inches(1),
    text: 'raster test',
  });
  return { pres, slide };
};

// PNG magic signature: first 4 bytes are 89 50 4E 47.
const isPng = (bytes: Uint8Array): boolean =>
  bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47;

describe('renderSlideToRgba (Node)', () => {
  it('returns image with expected pixel dimensions and correct aspect ratio', async () => {
    const { pres, slide } = await buildTestSlide();
    const targetWidth = 320;
    const { image } = renderSlideToRgba(pres, slide, { width: targetWidth });

    expect(image.width).toBe(targetWidth);

    // The aspect ratio should match the slide (16:9 ≈ 1.777).
    const slideSize = getSlideSize(pres)!;
    const expectedHeight = Math.round((targetWidth * slideSize.height) / slideSize.width);
    // Allow ±1 pixel rounding difference between integer rounding approaches.
    expect(Math.abs(image.height - expectedHeight)).toBeLessThanOrEqual(1);

    // RGBA: 4 bytes per pixel.
    expect(image.data.length).toBe(image.width * image.height * 4);
  });

  it('the PNG output starts with the PNG signature bytes', async () => {
    const { pres, slide } = await buildTestSlide();
    const { png } = renderSlideToRgba(pres, slide, { width: 320 });
    expect(isPng(png)).toBe(true);
  });

  it('the RGBA buffer is not entirely white/blank', async () => {
    const { pres, slide } = await buildTestSlide();
    const { image } = renderSlideToRgba(pres, slide, { width: 320 });
    // Count pixels that differ from white (R=255, G=255, B=255).
    let nonWhite = 0;
    for (let i = 0; i < image.data.length; i += 4) {
      if (image.data[i] !== 255 || image.data[i + 1] !== 255 || image.data[i + 2] !== 255) {
        nonWhite++;
      }
    }
    expect(nonWhite).toBeGreaterThan(0);
  });

  it('render is byte-identical when called twice (determinism)', async () => {
    const { pres, slide } = await buildTestSlide();
    const opts = { width: 320 };
    const first = renderSlideToRgba(pres, slide, opts).png;
    const second = renderSlideToRgba(pres, slide, opts).png;
    // Compare via string encoding to get a useful diff on failure.
    expect(Buffer.from(first).toString('hex')).toBe(Buffer.from(second).toString('hex'));
  });
});

describe('renderSlideToImage (Node)', () => {
  it('returns a valid non-empty PNG', async () => {
    const { pres, slide } = await buildTestSlide();
    const png = renderSlideToImage(pres, slide, { width: 320 });
    expect(isPng(png)).toBe(true);
    expect(png.length).toBeGreaterThan(100);
  });

  it('output matches the png from renderSlideToRgba for the same input', async () => {
    // Both functions call the same internal rasterize(); their PNG outputs
    // should be identical for the same parameters.
    const { pres, slide } = await buildTestSlide();
    const opts = { width: 320 };
    const fromImage = renderSlideToImage(pres, slide, opts);
    const { png: fromRgba } = renderSlideToRgba(pres, slide, opts);
    expect(Buffer.from(fromImage).toString('hex')).toBe(Buffer.from(fromRgba).toString('hex'));
  });
});
