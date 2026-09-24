import { readFile } from 'node:fs/promises';
import { unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import {
  getShapeGradientFill,
  getShapeGradientFillEffective,
  getSlideShapes,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeGradientFill,
} from '../src/api/index.ts';
import { renderSlideSvg } from '../packages/preview/src/render-slide.ts';

const transformedGradient = async (path: 'linear' | 'circle') => {
  const pres = await loadPresentation(
    await readFile(new URL('./fixtures/minimal/one-text-slide.pptx', import.meta.url)),
  );
  setShapeGradientFill(getSlideShapes(getSlides(pres)[0]!)[0]!, {
    path,
    stops: [
      { offset: 0, color: '#000000' },
      { offset: 1, color: '#FFFFFF' },
    ],
  });
  const parts = unzipSync(await savePresentation(pres));
  const name = 'ppt/slides/slide1.xml';
  const xml = new TextDecoder().decode(parts[name]);
  parts[name] = new TextEncoder().encode(
    xml.replace(
      /<a:gs pos="0">.*?<\/a:gs>/,
      '<a:gs pos="0"><a:schemeClr val="tx1"><a:lumMod val="0"/><a:lumOff val="50000"/><a:alpha val="50000"/><a:alphaMod val="50000"/></a:schemeClr></a:gs>',
    ),
  );
  return loadPresentation(zipSync(parts));
};

describe('imported gradient stop color transforms', () => {
  it('retains the theme token and exposes resolved brightness and opacity', async () => {
    const pres = await transformedGradient('linear');
    const shape = getSlideShapes(getSlides(pres)[0]!)[0]!;
    expect(getShapeGradientFill(shape)?.stops[0]).toMatchObject({
      color: 'scheme:tx1',
      opacity: 0.25,
    });
    expect(getShapeGradientFillEffective(pres, shape)?.stops[0]).toMatchObject({
      color: 'scheme:tx1',
      resolvedColor: '#808080',
      opacity: 0.25,
    });
    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(
      getShapeGradientFillEffective(reloaded, getSlideShapes(getSlides(reloaded)[0]!)[0]!),
    ).toEqual(getShapeGradientFillEffective(pres, shape));
  });

  it.each(['linear', 'circle'] as const)(
    'renders %s gradient stop brightness and opacity',
    async (path) => {
      const pres = await transformedGradient(path);
      const svg = renderSlideSvg(pres, getSlides(pres)[0]!);
      expect(svg).toMatch(/<stop[^>]*stop-color="#808080"[^>]*stop-opacity="0.25"/);
    },
  );
});
