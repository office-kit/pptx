// getShapeFillOpacity / getShapeStrokeOpacity — the alpha channel OOXML keeps
// beside the color. No authoring API writes `<a:alpha>`, so the element is
// injected at the OPC zip layer, the same way the preview tests do it.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideShape,
  findSlideLayout,
  getShapeFillColorResolved,
  getShapeFillOpacity,
  getShapeStrokeColorResolved,
  getShapeStrokeOpacity,
  getSlides,
  getSlideShapes,
  inches,
  loadPresentation,
  savePresentation,
  setShapeFill,
  setShapeFillOpacity,
  setShapeStrokeOpacity,
  setShapeStroke,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));
const dec = new TextDecoder();
const enc = new TextEncoder();

// Adds one rect (fill #3366CC, 1pt #000000 outline) to a fresh slide, then
// rewrites the saved slide XML so both colors carry the given alpha children.
const loadRectWithAlpha = async (fillAlpha: string, strokeAlpha: string) => {
  const pres = await loadPresentation(await readFile(fixturePath));
  const layout = findSlideLayout(pres, 'Blank');
  if (!layout) throw new Error('Blank layout missing');
  const slide = addSlide(pres, { layout });
  const rect = addSlideShape(slide, {
    preset: 'rect',
    x: inches(1),
    y: inches(1),
    w: inches(2),
    h: inches(2),
  });
  setShapeFill(rect, '#3366CC');
  setShapeStroke(rect, { color: '#000000', widthEmu: 12_700 });
  const { entries } = readZip(await savePresentation(pres));
  const modified = entries.map((e) => {
    if (!/slides\/slide\d+\.xml$/.test(e.name)) return e;
    const xml = dec
      .decode(e.data)
      .replace('<a:srgbClr val="3366CC"/>', `<a:srgbClr val="3366CC">${fillAlpha}</a:srgbClr>`)
      .replace('<a:srgbClr val="000000"/>', `<a:srgbClr val="000000">${strokeAlpha}</a:srgbClr>`);
    return { name: e.name, data: enc.encode(xml) };
  });
  const reloaded = await loadPresentation(writeZip(modified));
  const last = getSlides(reloaded).at(-1);
  if (!last) throw new Error('slide missing');
  const shape = getSlideShapes(last).at(-1);
  if (!shape) throw new Error('rect missing');
  return { pres: reloaded, shape };
};

describe('fn API: getShapeFillOpacity / getShapeStrokeOpacity', () => {
  it('returns null for opaque fills and outlines', async () => {
    const { shape } = await loadRectWithAlpha('', '');
    expect(getShapeFillOpacity(shape)).toBeNull();
    expect(getShapeStrokeOpacity(shape)).toBeNull();
  });

  it('reads the alpha of the fill and the outline independently', async () => {
    const { pres, shape } = await loadRectWithAlpha(
      '<a:alpha val="27000"/>',
      '<a:alpha val="50000"/>',
    );
    expect(getShapeFillOpacity(shape)).toBeCloseTo(0.27, 6);
    expect(getShapeStrokeOpacity(shape)).toBeCloseTo(0.5, 6);
    // The color readers stay alpha-free: OOXML encodes the two separately.
    expect(getShapeFillColorResolved(pres, shape)).toBe('#3366CC');
    expect(getShapeStrokeColorResolved(pres, shape)).toBe('#000000');
  });
});

it('width-only outline changes preserve color transforms through export', async () => {
  const { pres, shape } = await loadRectWithAlpha('', '<a:alpha val="35000"/>');
  setShapeStroke(shape, { widthEmu: 38100 });
  expect(getShapeStrokeColorResolved(pres, shape)).toBe('#000000');
  expect(getShapeStrokeOpacity(shape)).toBe(0.35);
  const reloaded = await loadPresentation(await savePresentation(pres));
  const saved = getSlideShapes(getSlides(reloaded).at(-1)!).at(-1)!;
  expect(getShapeStrokeOpacity(saved)).toBe(0.35);
  expect(getShapeStrokeColorResolved(reloaded, saved)).toBe('#000000');
});

it('absolute opacity replaces alpha transforms, survives export and rejects invalid values', async () => {
  const { pres, shape } = await loadRectWithAlpha(
    '<a:alpha val="80000"/><a:alphaMod val="50000"/>',
    '<a:alpha val="60000"/><a:alphaOff val="10000"/>',
  );
  setShapeFillOpacity(shape, 0.25);
  setShapeStrokeOpacity(shape, 0);
  expect(getShapeFillOpacity(shape)).toBe(0.25);
  expect(getShapeStrokeOpacity(shape)).toBe(0);
  for (const value of [-1, 1.1, NaN, Infinity]) {
    expect(() => setShapeFillOpacity(shape, value)).toThrow();
    expect(() => setShapeStrokeOpacity(shape, value)).toThrow();
  }
  expect(getShapeFillOpacity(shape)).toBe(0.25);
  const loaded = await loadPresentation(await savePresentation(pres));
  const saved = getSlideShapes(getSlides(loaded).at(-1)!).at(-1)!;
  expect(getShapeFillOpacity(saved)).toBe(0.25);
  expect(getShapeStrokeOpacity(saved)).toBe(0);
  expect(getShapeFillColorResolved(loaded, saved)).toBe('#3366CC');
  setShapeFillOpacity(saved, 1);
  expect(getShapeFillOpacity(saved)).toBe(1);
});
