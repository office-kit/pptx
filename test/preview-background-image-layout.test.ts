import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import {
  addBlankSlide,
  applySlideBackgroundToAll,
  getSlideBackgroundImageIntrinsicSize,
  getSlides,
  getSlideSize,
  inches,
  loadPresentation,
  pt,
  savePresentation,
  setSlideBackgroundImage,
  setSlideBackgroundImageFillLayout,
  setSlideBackgroundImageOpacity,
  type ImageTileFlip,
} from '../src/api/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';
import { renderSlideToRgba } from '../packages/preview/src/node.ts';
import { buildPng } from './lib/build-png.ts';

async function fixture() {
  const pres = await loadPresentation(
    await readFile(new URL('./fixtures/minimal/blank.pptx', import.meta.url)),
  );
  const slide = addBlankSlide(pres);
  setSlideBackgroundImage(
    slide,
    buildPng(24, 24, (x, y) => [x < 12 ? 255 : 0, y < 12 ? 255 : 0, 0]),
  );
  return { pres, slide };
}

function pixels(
  pres: Parameters<typeof renderSlideToRgba>[0],
  slide: Parameters<typeof renderSlideToRgba>[1],
) {
  const { image } = renderSlideToRgba(pres, slide, {
    width: Math.round(getSlideSize(pres)!.width / 9525),
  });
  return (x: number, y: number) =>
    Array.from(image.data.slice((y * image.width + x) * 4, (y * image.width + x) * 4 + 3));
}

it('stretches the entire source into background offsets instead of cropping it to cover', async () => {
  const { pres, slide } = await fixture();
  setSlideBackgroundImageFillLayout(slide, {
    mode: 'stretch',
    left: 0.25,
    top: 0.25,
    right: 0.25,
    bottom: 0.25,
  });
  const size = getSlideSize(pres)!;
  const w = Math.round(size.width / 9525),
    h = Math.round(size.height / 9525);
  const pixel = pixels(pres, slide);
  expect(pixel(10, 10)).toEqual([255, 255, 255]);
  expect(pixel(Math.round(w * 0.3), Math.round(h * 0.3))).toEqual([255, 255, 0]);
  expect(pixel(Math.round(w * 0.7), Math.round(h * 0.3))).toEqual([0, 255, 0]);
  expect(pixel(Math.round(w * 0.3), Math.round(h * 0.7))).toEqual([255, 0, 0]);
  expect(pixel(Math.round(w * 0.7), Math.round(h * 0.7))).toEqual([0, 0, 0]);
  setSlideBackgroundImageFillLayout(slide, { mode: 'stretch', left: -0.5 });
  expect(renderSlideToSvg(pres, slide)).toContain(`x="${(-w / 2).toFixed(2)}"`);
});

it.each(['none', 'x', 'y', 'xy'] as ImageTileFlip[])(
  'tiles inherited backgrounds with %s reflection and preserves the master after an override',
  async (flip) => {
    const { pres, slide } = await fixture();
    setSlideBackgroundImageFillLayout(slide, { mode: 'tile', flip });
    const sibling = addBlankSlide(pres);
    applySlideBackgroundToAll(pres, slide);
    expect(getSlideBackgroundImageIntrinsicSize(slide)).toEqual({
      width: inches(0.25),
      height: inches(0.25),
    });
    const pixel = pixels(pres, slide);
    expect(pixel(6, 6)).toEqual([255, 255, 0]);
    expect(pixel(30, 6)).toEqual([flip === 'x' || flip === 'xy' ? 0 : 255, 255, 0]);
    expect(pixel(6, 30)).toEqual([255, flip === 'y' || flip === 'xy' ? 0 : 255, 0]);
    setSlideBackgroundImageOpacity(slide, 0);
    expect(pixels(pres, slide)(6, 6)).toEqual([255, 255, 255]);
    expect(pixels(pres, sibling)(6, 6)).toEqual([255, 255, 0]);
  },
);

it('persists background tile scale, alignment and offsets and renders zero-size tiles empty', async () => {
  const { pres, slide } = await fixture();
  setSlideBackgroundImageFillLayout(slide, {
    mode: 'tile',
    scaleX: 0.5,
    scaleY: 0.25,
    alignment: 'br',
    offsetX: pt(6),
    offsetY: pt(-3),
  });
  const saved = await loadPresentation(await savePresentation(pres));
  const restored = getSlides(saved).at(-1)!;
  const size = getSlideSize(saved)!;
  const w = size.width / 9525,
    h = size.height / 9525;
  const svg = renderSlideToSvg(saved, restored);
  expect(svg).toContain(
    `patternUnits="userSpaceOnUse" x="${(w - 12 + 8).toFixed(2)}" y="${(h - 6 - 4).toFixed(2)}" width="12.00" height="6.00"`,
  );
  setSlideBackgroundImageFillLayout(restored, { mode: 'tile', scaleX: 0 });
  expect(pixels(saved, restored)(6, 6)).toEqual([255, 255, 255]);
});
