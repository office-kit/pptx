import { expect, it } from 'vitest';
import * as pptx from '../src/api/index.ts';
import { renderSlideToRgba } from '../packages/preview/src/node.ts';
import { buildPng } from './lib/build-png.ts';

it.each([
  ['fill', 4, 2, { left: 0.25, right: 0.25, top: 0, bottom: 0 }],
  ['fit', 4, 2, { left: 0, right: 0, top: -0.5, bottom: -0.5 }],
  ['fill', 2, 4, { left: 0, right: 0, top: 0.25, bottom: 0.25 }],
  ['fit', 2, 4, { left: -0.5, right: -0.5, top: 0, bottom: 0 }],
] as const)('%s for %s × %s preserves frame and source proportions', async (mode, w, h, crop) => {
  const p = pptx.createPresentation();
  const slide = pptx.addBlankSlide(p);
  const bg = pptx.addSlideShape(slide, {
    preset: 'rect',
    x: pptx.emu(0),
    y: pptx.emu(0),
    w: pptx.inches(5),
    h: pptx.inches(5),
  });
  pptx.setShapeFill(bg, '#0000FF');
  const bounds = { x: pptx.inches(1), y: pptx.inches(1), w: pptx.inches(2), h: pptx.inches(2) };
  const image = pptx.addSlideImage(slide, buildPng(w, h, [0, 255, 0]), bounds);
  pptx.setShapeImageCrop(image, { left: 0.1, top: 0.1 });
  pptx.setShapeImageFit(image, mode);
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  const savedSlide = pptx.getSlides(saved)[0]!;
  const savedImage = pptx.getSlideShapes(savedSlide)[1]!;
  expect(pptx.getShapeBounds(savedImage)).toEqual(bounds);
  expect(pptx.getShapeImageCrop(savedImage)).toEqual(crop);
  const { image: raster } = renderSlideToRgba(saved, savedSlide);
  const pixel = (x: number, y: number) =>
    Array.from(raster.data.slice((y * raster.width + x) * 4, (y * raster.width + x) * 4 + 4));
  expect(pixel(192, 192)).toEqual([0, 255, 0, 255]);
  expect(pixel(w > h ? 192 : 100, w > h ? 100 : 192)).toEqual(
    mode === 'fill' ? [0, 255, 0, 255] : [0, 0, 255, 255],
  );
  expect(pixel(w > h ? 192 : 150, w > h ? 150 : 192)).toEqual([0, 255, 0, 255]);
});

it('rejects invalid crop without altering the saved previous crop', async () => {
  const p = pptx.createPresentation(),
    slide = pptx.addBlankSlide(p);
  const image = pptx.addSlideImage(slide, buildPng(4, 2, [0, 255, 0]), {
    x: pptx.emu(0),
    y: pptx.emu(0),
    w: pptx.inches(2),
    h: pptx.inches(2),
  });
  pptx.setShapeImageFit(image, 'fit');
  const before = pptx.getSlideXmlString(slide);
  for (const crop of [
    { left: 0.6, right: 0.5 },
    { top: NaN },
    { bottom: Infinity },
    { left: -30000 },
  ]) {
    expect(() => pptx.setShapeImageCrop(image, crop)).toThrow();
    expect(pptx.getSlideXmlString(slide)).toBe(before);
  }
});
