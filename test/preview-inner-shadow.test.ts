import { readFile } from 'node:fs/promises';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideShape,
  getSlides,
  getSlideSize,
  inches,
  loadPresentation,
  savePresentation,
  setShapeFill,
  setShapeGlow,
  setShapeShadow,
} from '../src/api/index.ts';
import { renderSlideToRgba } from '../packages/preview/src/node.ts';

it.each([false, true])('paints an inner shadow over an opaque fill (glow: %s)', async (glow) => {
  const pres = await loadPresentation(
    await readFile(new URL('./fixtures/minimal/blank.pptx', import.meta.url)),
  );
  const slide = addBlankSlide(pres);
  const shape = addSlideShape(slide, {
    preset: 'rect',
    x: inches(1),
    y: inches(1),
    w: inches(2),
    h: inches(2),
  });
  setShapeFill(shape, '#FF0000');
  setShapeShadow(shape, {
    color: '#000000',
    blurEmu: 0,
    offsetEmu: inches(0.2),
    angleDeg: 0,
    opacity: 0.5,
  });
  if (glow) setShapeGlow(shape, { color: '#00FF00', radiusEmu: inches(0.1) });
  const zip = unzipSync(await savePresentation(pres));
  const part = Object.keys(zip).find(
    (name) =>
      /^ppt\/slides\/slide\d+\.xml$/.test(name) && strFromU8(zip[name]!).includes('outerShdw'),
  )!;
  zip[part] = strToU8(strFromU8(zip[part]!).replaceAll('a:outerShdw', 'a:innerShdw'));
  const imported = await loadPresentation(zipSync(zip));
  const target = getSlides(imported).at(-1)!;
  const { image } = renderSlideToRgba(imported, target, {
    width: Math.round(getSlideSize(imported)!.width / 9525),
  });
  const pixel = (x: number, y: number) =>
    Array.from(image.data.slice((y * image.width + x) * 4, (y * image.width + x) * 4 + 3));
  // At 96 dpi the rectangle spans 96..288 and the inset spans 19 pixels.
  const inset = pixel(104, 192);
  // SVG filters composite in linear RGB: half red maps to about 188 sRGB.
  expect(inset[0]).toBeGreaterThan(175);
  expect(inset[0]).toBeLessThan(200);
  expect(inset.slice(1)).toEqual([0, 0]);
  expect(pixel(192, 192)).toEqual([255, 0, 0]);
  expect(pixel(70, 192)).toEqual([255, 255, 255]);
});
