// Linear gradient fill on a shape.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getSlideShapes,
  getSlideXmlString,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeGradientFill,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const slideXml = async (bytes: Uint8Array, slideIndex: number): Promise<string> => {
  const pres = await loadPresentation(bytes);
  return getSlideXmlString(getSlides(pres)[slideIndex]!);
};

describe('fn API: setShapeGradientFill', () => {
  it('writes a gradFill with the configured stops and angle', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    setShapeGradientFill(shape, {
      stops: [
        { offset: 0, color: '#FF0000' },
        { offset: 1, color: '#0000FF' },
      ],
      angleDeg: 90,
    });
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('<a:gradFill');
    expect(xml).toContain('<a:gs pos="0">');
    expect(xml).toContain('<a:gs pos="100000">');
    expect(xml).toContain('FF0000');
    expect(xml).toContain('0000FF');
    // 90 degrees × 60000 = 5400000
    expect(xml).toContain('ang="5400000"');
  });

  it('default angle is 90 degrees when omitted', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    setShapeGradientFill(shape, {
      stops: [
        { offset: 0, color: '#FFFFFF' },
        { offset: 1, color: '#000000' },
      ],
    });
    expect(await slideXml(await savePresentation(pres), 0)).toContain('ang="5400000"');
  });

  it('supports multi-stop gradients', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    setShapeGradientFill(shape, {
      stops: [
        { offset: 0, color: '#FF0000' },
        { offset: 0.5, color: '#00FF00' },
        { offset: 1, color: '#0000FF' },
      ],
    });
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('<a:gs pos="0">');
    expect(xml).toContain('<a:gs pos="50000">');
    expect(xml).toContain('<a:gs pos="100000">');
  });

  it('rejects fewer than two stops', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    expect(() => setShapeGradientFill(shape, { stops: [{ offset: 0, color: '#FFFFFF' }] })).toThrow(
      /two stops/,
    );
  });

  it('rejects offsets outside [0, 1]', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    expect(() =>
      setShapeGradientFill(shape, {
        stops: [
          { offset: 0, color: '#FFFFFF' },
          { offset: 1.5, color: '#000000' },
        ],
      }),
    ).toThrow(RangeError);
  });
});
