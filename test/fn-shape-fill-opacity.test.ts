// Imported color transforms and authored opacity must round-trip independently.

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

describe('solid paint opacity editing', () => {
  it('replaces alpha transforms while preserving color transforms and line width', async () => {
    const { pres, shape } = await loadRectWithAlpha(
      '<a:alpha val="60000"/><a:alphaMod val="50000"/><a:tint val="10000"/>',
      '<a:alpha val="50000"/><a:alphaOff val="10000"/>',
    );
    const fillColor = getShapeFillColorResolved(pres, shape);
    setShapeFill(shape, { opacity: 0.25 });
    setShapeStroke(shape, { opacity: 0.75 });
    expect(getShapeFillColorResolved(pres, shape)).toBe(fillColor);
    expect(getShapeFillOpacity(shape)).toBe(0.25);
    expect(getShapeStrokeOpacity(shape)).toBe(0.75);
    const bytes = await savePresentation(pres);
    const xml = readZip(bytes)
      .entries.filter((entry) => /slides\/slide\d+\.xml$/.test(entry.name))
      .map((entry) => dec.decode(entry.data))
      .join('');
    expect(xml).toContain('<a:tint val="10000"/>');
    expect(xml).not.toContain('<a:alphaMod');
    expect(xml).not.toContain('<a:alphaOff');
    expect(xml).toContain('w="12700"');
    const loaded = await loadPresentation(bytes);
    const restored = getSlideShapes(getSlides(loaded).at(-1)!).at(-1)!;
    expect(getShapeFillOpacity(restored)).toBe(0.25);
    expect(getShapeStrokeOpacity(restored)).toBe(0.75);
  });

  it('keeps opacity when editing colors with options and accepts transparent and opaque endpoints', async () => {
    const { shape } = await loadRectWithAlpha('<a:alpha val="27000"/>', '<a:alpha val="50000"/>');
    setShapeFill(shape, { color: 'accent1' });
    setShapeStroke(shape, { color: 'accent2' });
    expect(getShapeFillOpacity(shape)).toBe(0.27);
    expect(getShapeStrokeOpacity(shape)).toBe(0.5);
    setShapeFill(shape, { opacity: 0 });
    setShapeStroke(shape, { opacity: 1 });
    expect(getShapeFillOpacity(shape)).toBe(0);
    expect(getShapeStrokeOpacity(shape)).toBe(1);
    setShapeFill(shape, '#FFFFFF');
    expect(getShapeFillOpacity(shape)).toBeNull();
  });

  it('rejects invalid opacity without changing the paint', async () => {
    const { shape } = await loadRectWithAlpha('<a:alpha val="27000"/>', '<a:alpha val="50000"/>');
    for (const opacity of [-0.1, 1.1, NaN, Infinity]) {
      expect(() => setShapeFill(shape, { opacity })).toThrow(RangeError);
      expect(() => setShapeStroke(shape, { opacity })).toThrow(RangeError);
      expect(getShapeFillOpacity(shape)).toBe(0.27);
      expect(getShapeStrokeOpacity(shape)).toBe(0.5);
    }
  });
});
