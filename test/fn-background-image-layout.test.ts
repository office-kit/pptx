import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import {
  addBlankSlide,
  applySlideBackgroundToAll,
  getMediaParts,
  getShapeImageBytes,
  getShapeKind,
  getSlideBackground,
  getSlideBackgroundImageBytes,
  getSlideBackgroundImageFillLayout,
  getSlideShapes,
  getSlides,
  getSlideXmlString,
  loadPresentation,
  pt,
  savePresentation,
  setSlideBackgroundImage,
  setSlideBackgroundImageFillLayout,
} from '../src/api/index.ts';

it('edits inherited image placement without replacing media or changing the master', async () => {
  const pres = await loadPresentation(
    await readFile(new URL('./fixtures/minimal/one-image-slide.pptx', import.meta.url)),
  );
  const slide = getSlides(pres)[0]!;
  const picture = getSlideShapes(slide).find((shape) => getShapeKind(shape) === 'picture')!;
  const bytes = getShapeImageBytes(picture)!;
  setSlideBackgroundImage(slide, bytes);
  const other = addBlankSlide(pres);
  const stretch = {
    mode: 'stretch' as const,
    left: 0.25,
    right: -0.1,
    top: 0,
    bottom: 0.2,
    rotateWithShape: false,
  };
  setSlideBackgroundImageFillLayout(slide, stretch);
  applySlideBackgroundToAll(pres, slide);
  expect(getSlideBackground(slide).kind).toBe('inherit');
  expect(getSlideBackgroundImageFillLayout(slide)).toEqual(stretch);
  const mediaCount = getMediaParts(pres).length;
  const tile = {
    mode: 'tile' as const,
    offsetX: pt(12),
    offsetY: pt(-6),
    scaleX: 0.5,
    scaleY: 0.75,
    alignment: 'ctr' as const,
    flip: 'xy' as const,
    rotateWithShape: false,
  };
  setSlideBackgroundImageFillLayout(slide, tile);
  expect(getSlideBackgroundImageBytes(slide)).toEqual(bytes);
  expect(getMediaParts(pres)).toHaveLength(mediaCount);
  expect(getSlideBackgroundImageFillLayout(slide)).toEqual(tile);
  expect(getSlideBackgroundImageFillLayout(other)).toEqual(stretch);
  const saved = await loadPresentation(await savePresentation(pres));
  expect(getSlideBackgroundImageFillLayout(getSlides(saved)[0]!)).toEqual(tile);
  expect(getSlideBackgroundImageBytes(getSlides(saved)[0]!)).toEqual(bytes);
  expect(getSlideBackgroundImageFillLayout(getSlides(saved)[1]!)).toEqual(stretch);
  const before = getSlideXmlString(slide);
  expect(() => setSlideBackgroundImageFillLayout(slide, { mode: 'stretch', left: NaN })).toThrow();
  expect(getSlideXmlString(slide)).toBe(before);
});
