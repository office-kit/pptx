import { readFile } from 'node:fs/promises';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideImage,
  createPresentation,
  getShapeImageDuotone,
  getShapeKind,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';
import { renderSlideToRgba } from '../packages/preview/src/node.ts';
import { buildPng } from './lib/build-png.ts';

it.each([
  ['<a:scrgbClr r="50%" g="50%" b="50%"/>', '#BCBCBC'],
  ['<a:hslClr hue="0" sat="100%" lum="50%"><a:shade val="50%"/></a:hslClr>', '#BC0000'],
])('retains alternate duotone colors and their order: %s', async (color, expected) => {
  const parts = unzipSync(
    await readFile(new URL('./fixtures/minimal/one-image-slide.pptx', import.meta.url)),
  );
  const slidePart = 'ppt/slides/slide1.xml';
  const xml = strFromU8(parts[slidePart]!).replace(
    '<a:blip r:embed="rId2"/>',
    `<a:blip r:embed="rId2"><a:duotone>${color}<a:srgbClr val="123456"/></a:duotone></a:blip>`,
  );
  expect(xml).toContain(`<a:duotone>${color}`);
  parts[slidePart] = strToU8(xml);
  const imported = await loadPresentation(zipSync(parts));
  const reloaded = await loadPresentation(await savePresentation(imported));
  for (const pres of [imported, reloaded]) {
    const picture = getSlideShapes(getSlides(pres)[0]!).find(
      (shape) => getShapeKind(shape) === 'picture',
    )!;
    expect(getShapeImageDuotone(pres, picture)).toEqual({
      firstColor: expected,
      secondColor: '#123456',
    });
  }
});

it.each([
  [
    '<a:scrgbClr r="50%" g="50%" b="50%"/><a:hslClr hue="0" sat="100%" lum="50%"/>',
    [188, 188, 188],
    [255, 0, 0],
  ],
  ['<a:srgbClr val="010509"/><a:srgbClr val="CAB292"/>', [1, 5, 9], [202, 178, 146]],
])('renders exact duotone endpoints after save and reload: %s', async (colors, first, second) => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  addSlideImage(
    slide,
    buildPng(600, 200, (x) => (x < 200 ? [0, 0, 0] : x < 400 ? [128, 128, 128] : [255, 255, 255])),
    {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
    },
  );
  const parts = unzipSync(await savePresentation(pres));
  const slidePart = 'ppt/slides/slide1.xml';
  const xml = strFromU8(parts[slidePart]!).replace(
    /<a:blip\b([^>]*)\/>/,
    `<a:blip$1><a:duotone>${colors}</a:duotone></a:blip>`,
  );
  expect(xml).toContain('<a:duotone>');
  parts[slidePart] = strToU8(xml);
  const imported = await loadPresentation(zipSync(parts));
  const reloaded = await loadPresentation(await savePresentation(imported));
  const { image } = renderSlideToRgba(reloaded, getSlides(reloaded)[0]!, { width: 960 });
  // Samples are inside the one-inch-high picture, away from interpolation edges.
  const pixel = (x: number) =>
    Array.from(image.data.slice((20 * image.width + x) * 4, (20 * image.width + x) * 4 + 3));
  expect(pixel(20)).toEqual(first);
  expect(pixel(120)).toEqual(second);
  const midpoint = pixel(70);
  for (let channel = 0; channel < 3; channel++) {
    const expected = first[channel]! + ((second[channel]! - first[channel]!) * 128) / 255;
    expect(Math.abs(midpoint[channel]! - expected)).toBeLessThanOrEqual(1);
  }
});
