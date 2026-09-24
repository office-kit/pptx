import { expect, it } from 'vitest';
import * as pptx from '../src/api/index.ts';
import { buildPng } from './lib/build-png.ts';

it.each([1, 2 / 3, 16 / 9])(
  'crops to %s around the same center without scaling the image',
  async (ratio) => {
    const p = pptx.createPresentation(),
      s = pptx.addBlankSlide(p);
    const box = {
      x: pptx.emu(100000),
      y: pptx.emu(200000),
      w: pptx.emu(4000000),
      h: pptx.emu(2000000),
    };
    const bytes = buildPng(4, 2, [0, 255, 0]);
    const pic = pptx.addSlideImage(s, bytes, box);
    pptx.setShapeRotation(pic, 35);
    pptx.setShapeImageCrop(pic, { left: 0.1, right: 0.2, top: -0.2, bottom: 0.1 });
    pptx.setShapeImageCropAspectRatio(pic, ratio);
    const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
    const restored = pptx.getSlideShapes(pptx.getSlides(saved)[0]!)[0]!;
    const b = pptx.getShapeBounds(restored)!,
      c = pptx.getShapeImageCrop(restored)!;
    expect(b.w / b.h).toBeCloseTo(ratio, 5);
    // OOXML integer coordinates can move the center by half an EMU.
    expect(Math.abs(b.x + b.w / 2 - (box.x + box.w / 2))).toBeLessThanOrEqual(0.5);
    expect(Math.abs(b.y + b.h / 2 - (box.y + box.h / 2))).toBeLessThanOrEqual(0.5);
    // Each crop edge is quantized to 1/100000 of the source image.
    expect(Math.abs(1 - c.left! - c.right! - (b.w / box.w) * 0.7)).toBeLessThanOrEqual(0.00001);
    expect(Math.abs(1 - c.top! - c.bottom! - (b.h / box.h) * 1.1)).toBeLessThanOrEqual(0.00001);
    expect(c.right! - c.left!).toBeCloseTo(0.1);
    expect(pptx.getShapeRotation(restored)).toBe(35);
    expect(Array.from(pptx.getShapeImageBytes(restored)!)).toEqual(Array.from(bytes));
  },
);

it('invalid ratios leave the picture XML unchanged', () => {
  const p = pptx.createPresentation(),
    s = pptx.addBlankSlide(p);
  const pic = pptx.addSlideImage(s, buildPng(2, 1, [0, 255, 0]), {
    x: pptx.emu(0),
    y: pptx.emu(0),
    w: pptx.emu(10000),
    h: pptx.emu(5000),
  });
  const before = pptx.getSlideXmlString(s);
  for (const ratio of [0, -1, NaN, Infinity, 1e20, 1e-20]) {
    expect(() => pptx.setShapeImageCropAspectRatio(pic, ratio)).toThrow();
    expect(pptx.getSlideXmlString(s)).toBe(before);
  }
});
