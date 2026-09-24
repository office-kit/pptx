import { expect, it } from 'vitest';
import * as pptx from '../src/api/index.ts';
import { renderSlideToRgba } from '../packages/preview/src/node.ts';
import { buildPng } from './lib/build-png.ts';

it('picture contrast preserves neutral colors and expands or compresses tonal differences', async () => {
  const p = pptx.createPresentation();
  const slide = pptx.addBlankSlide(p);
  const picture = pptx.addSlideImage(slide, buildPng(2, 2, [64, 128, 192]), {
    x: pptx.emu(0),
    y: pptx.emu(0),
    w: pptx.inches(1),
    h: pptx.inches(1),
  });
  const sample = async (contrast: number, brightness = 0) => {
    pptx.setShapeImageContrast(picture, contrast);
    pptx.setShapeImageBrightness(picture, brightness);
    const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
    const { image } = renderSlideToRgba(saved, pptx.getSlides(saved)[0]!);
    const offset = (48 * image.width + 48) * 4;
    return Array.from(image.data.slice(offset, offset + 4));
  };
  const normal = await sample(0);
  expect(normal).toEqual([64, 128, 192, 255]);
  const increased = await sample(0.5);
  expect(increased[0]).toBeLessThan(normal[0]!);
  expect(increased[2]).toBeGreaterThan(normal[2]!);
  expect(Math.abs(increased[1]! - normal[1]!)).toBeLessThanOrEqual(1);
  const decreased = await sample(-0.5);
  expect(decreased[0]).toBeGreaterThan(normal[0]!);
  expect(decreased[2]).toBeLessThan(normal[2]!);
  const flat = await sample(-1);
  expect(Math.max(...flat.slice(0, 3)) - Math.min(...flat.slice(0, 3))).toBe(0);
  const lighter = await sample(0, 0.2);
  const darker = await sample(0, -0.2);
  for (let i = 0; i < 3; i++) {
    expect(lighter[i]).toBeGreaterThan(normal[i]!);
    expect(darker[i]).toBeLessThan(normal[i]!);
  }
});
