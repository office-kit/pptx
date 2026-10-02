import { unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import {
  getShapeGradientFill,
  getSlideShapes,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeGradientFill,
} from '../src/api/index.ts';
import { readFile } from 'node:fs/promises';

const injectGradientStops = (xml: string): string =>
  xml.replace(
    /<a:gsLst>.*?<\/a:gsLst>/,
    '<a:gsLst>' +
      '<a:gs pos="0"><a:prstClr val="red"/></a:gs>' +
      '<a:gs pos="50000"><a:sysClr val="windowText" lastClr="123456"/></a:gs>' +
      '<a:gs pos="100000"><a:srgbClr val="ABCDEF"/></a:gs>' +
      '</a:gsLst>',
  );

describe('imported gradient preset colors', () => {
  it('keeps preset, system, and sRGB stops through save and reload', async () => {
    const presentation = await loadPresentation(
      await readFile(new URL('./fixtures/minimal/one-text-slide.pptx', import.meta.url)),
    );
    const shape = getSlideShapes(getSlides(presentation)[0]!)[0]!;
    setShapeGradientFill(shape, {
      stops: [
        { offset: 0, color: '#000000' },
        { offset: 1, color: '#FFFFFF' },
      ],
    });
    const authored = unzipSync(await savePresentation(presentation));
    const slideName = 'ppt/slides/slide1.xml';
    authored[slideName] = new TextEncoder().encode(
      injectGradientStops(new TextDecoder().decode(authored[slideName])),
    );

    const imported = await loadPresentation(zipSync(authored));
    const importedShape = getSlideShapes(getSlides(imported)[0]!)[0]!;
    expect(getShapeGradientFill(importedShape)?.stops.map((stop) => stop.color)).toEqual([
      '#FF0000',
      '#123456',
      '#ABCDEF',
    ]);

    const reloaded = await loadPresentation(await savePresentation(imported));
    const reloadedShape = getSlideShapes(getSlides(reloaded)[0]!)[0]!;
    expect(getShapeGradientFill(reloadedShape)?.stops.map((stop) => stop.color)).toEqual([
      '#FF0000',
      '#123456',
      '#ABCDEF',
    ]);
  });
});
