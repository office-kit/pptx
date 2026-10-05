import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import {
  getSlides,
  getSlideShapes,
  getShapePreset,
  setShapePreset,
  setShapeStroke,
  setShapeStrokeDash,
  setShapeNoStroke,
  setShapeRotation,
  getShapeBounds,
  getShapeImageBytes,
  getShapeImageCrop,
  setShapeImageCrop,
  setShapeAdjustValues,
  getShapeAdjustValues,
  getSlideXmlString,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

it('round-trips picture geometry without changing source bytes, crop or bounds', async () => {
  const pres = await loadPresentation(
    await readFile(new URL('./fixtures/minimal/one-image-slide.pptx', import.meta.url)),
  );
  const slide = getSlides(pres)[0]!;
  const picture = getSlideShapes(slide)[1]!;
  const bytes = getShapeImageBytes(picture);
  const bounds = getShapeBounds(picture);
  expect(getShapePreset(picture)).toBe('rect');
  setShapeImageCrop(picture, { left: 0.2, right: 0.1 });
  setShapePreset(picture, 'roundRect');
  setShapeAdjustValues(picture, { adj: 25000 });
  expect(renderSlideToSvg(pres, slide)).toMatch(/<clipPath[^>]*><rect[^>]*rx="/);
  setShapePreset(picture, 'ellipse');
  expect(getShapeAdjustValues(picture)).toEqual({});
  const saved = await loadPresentation(await savePresentation(pres));
  const restored = getSlideShapes(getSlides(saved)[0]!)[1]!;
  expect(getShapePreset(restored)).toBe('ellipse');
  expect(getShapeImageBytes(restored)).toEqual(bytes);
  expect(getShapeBounds(restored)).toEqual(bounds);
  expect(getShapeImageCrop(restored)).toEqual({ left: 0.2, right: 0.1, top: 0, bottom: 0 });
  expect(getSlideXmlString(getSlides(saved)[0]!)).toMatch(
    /<a:xfrm>[\s\S]*?<\/a:xfrm><a:prstGeom prst="ellipse"/,
  );
  const svg = renderSlideToSvg(saved, getSlides(saved)[0]!);
  expect(svg).toMatch(/<clipPath[^>]*><ellipse/);
  expect(svg).toMatch(/<g[^>]*clip-path="url\(#[^)]+\)"><image/);
  setShapePreset(restored, 'heart');
  expect(renderSlideToSvg(saved, getSlides(saved)[0]!)).toMatch(/<clipPath[^>]*><path/);
  setShapePreset(restored, 'hexagon');
  expect(renderSlideToSvg(saved, getSlides(saved)[0]!)).toMatch(/<clipPath[^>]*><polygon/);
  setShapePreset(restored, 'rect');
  setShapeImageCrop(restored, null);
  expect(renderSlideToSvg(saved, getSlides(saved)[0]!)).not.toContain('<clipPath');
});

it('changes native shape geometry and resets adjust guides without changing text or bounds', async () => {
  const pres = await loadPresentation(
    await readFile(new URL('./fixtures/minimal/one-text-slide.pptx', import.meta.url)),
  );
  const slide = getSlides(pres)[0]!;
  const shape = getSlideShapes(slide)[0]!;
  const bounds = getShapeBounds(shape);
  const text = getSlideXmlString(slide).match(/<a:t>.*?<\/a:t>/g);
  setShapePreset(shape, 'roundRect');
  setShapeAdjustValues(shape, { adj: 5000 });
  setShapePreset(shape, 'diamond');
  const restored = await loadPresentation(await savePresentation(pres));
  const result = getSlideShapes(getSlides(restored)[0]!)[0]!;
  expect(getShapePreset(result)).toBe('diamond');
  expect(getShapeBounds(result)).toEqual(bounds);
  expect(getShapeAdjustValues(result)).toEqual({});
  expect(getSlideXmlString(getSlides(restored)[0]!).match(/<a:t>.*?<\/a:t>/g)).toEqual(text);
});

it('renders picture outlines outside the crop clip with mask shape and line styling', async () => {
  const pres = await loadPresentation(
    await readFile(new URL('./fixtures/minimal/one-image-slide.pptx', import.meta.url)),
  );
  const slide = getSlides(pres)[0]!;
  const picture = getSlideShapes(slide)[1]!;
  setShapePreset(picture, 'ellipse');
  setShapeImageCrop(picture, { left: 0.2 });
  setShapeRotation(picture, 25);
  setShapeStroke(picture, { color: '#FF2200', widthEmu: 50800 });
  setShapeStrokeDash(picture, 'dash');
  const svg = renderSlideToSvg(pres, slide);
  expect(svg).toMatch(
    /<g transform="rotate\(25 [^"]+" fill="none" stroke="#FF2200" stroke-width="5.33" stroke-dasharray="[^"]+"><ellipse/,
  );
  expect(svg).toMatch(/<\/g><g transform="rotate\(25 [^"]+" fill="none"/);
  setShapeNoStroke(picture);
  expect(renderSlideToSvg(pres, slide)).not.toContain('stroke="#FF2200"');
});
