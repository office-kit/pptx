// `getShapeFill` introspection.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  clearShapeFill,
  getShapeFill,
  getSlides,
  inches,
  loadPresentation,
  setShapeFill,
  setShapeGradientFill,
  setShapeImageFill,
  setShapeNoFill,
  type SlideData,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const tinyPng = (): Uint8Array =>
  new Uint8Array([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x62, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
    0x42, 0x60, 0x82,
  ]);

const addRect = (slide: SlideData) =>
  addSlideShape(slide, {
    preset: 'rect',
    x: inches(0),
    y: inches(0),
    w: inches(2),
    h: inches(2),
  });

describe('fn API: getShapeFill', () => {
  it('returns inherit when spPr has no fill choice', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addRect(slide);
    // Fresh preset shape has no explicit fill on spPr.
    expect(getShapeFill(shape).kind).toBe('inherit');
  });

  it('reads back solid sRGB fill as #RRGGBB', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addRect(slide);
    setShapeFill(shape, '#ABCDEF');
    expect(getShapeFill(shape)).toEqual({ kind: 'solid', color: '#ABCDEF' });
  });

  it('reports gradient / image fill kinds', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const a = addRect(slide);
    const b = addRect(slide);

    setShapeGradientFill(a, {
      stops: [
        { offset: 0, color: '#FF0000' },
        { offset: 1, color: '#0000FF' },
      ],
    });
    setShapeImageFill(b, tinyPng(), { format: 'png' });

    expect(getShapeFill(a).kind).toBe('gradient');
    expect(getShapeFill(b).kind).toBe('image');
  });

  it('reports noFill', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addRect(slide);
    setShapeNoFill(shape);
    expect(getShapeFill(shape).kind).toBe('none');
  });

  it('clearShapeFill returns to inherit', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addRect(slide);
    setShapeFill(shape, '#112233');
    clearShapeFill(shape);
    expect(getShapeFill(shape).kind).toBe('inherit');
  });
});
