import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';
import {
  addTitleSlide,
  createPresentation,
  getShapeBodyPrEffective,
  getShapePlaceholderType,
  getSlideShapes,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeTextAutoFit,
  setShapeTextAnchorCenter,
  setShapeTextDirection,
} from '../src/api/index.ts';

describe('effective text body settings', () => {
  it('inherits autofit and direction, with explicit none and horizontal overrides', async () => {
    const original = createPresentation();
    addTitleSlide(original, 'Inherited text');
    const parts = unzipSync(await savePresentation(original));
    for (const [name, data] of Object.entries(parts)) {
      if (!/^ppt\/(slides|slideLayouts|slideMasters)\/[^/]+\.xml$/.test(name)) continue;
      const inherited = !name.startsWith('ppt/slides/');
      parts[name] = strToU8(
        strFromU8(data).replace(
          /<a:bodyPr\b[^>]*(?:\/>|>[\s\S]*?<\/a:bodyPr>)/g,
          inherited
            ? '<a:bodyPr vert="vert" anchorCtr="1"><a:normAutofit fontScale="65000" lnSpcReduction="10000"/></a:bodyPr>'
            : '<a:bodyPr/>',
        ),
      );
    }
    const pres = await loadPresentation(zipSync(parts));
    const shape = getSlideShapes(getSlides(pres)[0]!).find(
      (s) => getShapePlaceholderType(s) === 'ctrTitle',
    )!;
    expect(shape).toBeDefined();
    expect(getShapeBodyPrEffective(pres, shape)).toMatchObject({
      autoFit: 'normal',
      autoFitParams: { fontScale: 0.65, lnSpcReduction: 0.1 },
      vert: 'vert',
    });
    expect(getShapeBodyPrEffective(pres, shape).anchorCenter).toBe(true);
    setShapeTextAnchorCenter(shape, false);
    expect(getShapeBodyPrEffective(pres, shape).anchorCenter).toBe(false);
    const restored = await loadPresentation(await savePresentation(pres));
    const restoredShape = getSlideShapes(getSlides(restored)[0]!).find(
      (s) => getShapePlaceholderType(s) === 'ctrTitle',
    )!;
    expect(getShapeBodyPrEffective(restored, restoredShape).anchorCenter).toBe(false);
    setShapeTextAnchorCenter(shape, null);
    expect(getShapeBodyPrEffective(pres, shape).anchorCenter).toBe(true);
    setShapeTextAnchorCenter(shape, false);
    const slide = getSlides(pres)[0]!;
    const maxFont = (svg: string) =>
      Math.max(
        ...Array.from(svg.matchAll(/font-size(?:="|:)\s*([\d.]+)/g), (match) => Number(match[1])),
      );
    const inheritedSizes = (['svg', 'foreignObject'] as const).map((textLayout) =>
      maxFont(renderSlideToSvg(pres, slide, { textLayout })),
    );
    setShapeTextAutoFit(shape, 'none');
    for (const [index, textLayout] of (['svg', 'foreignObject'] as const).entries()) {
      const fullSize = maxFont(renderSlideToSvg(pres, slide, { textLayout }));
      expect(inheritedSizes[index]! / fullSize).toBeCloseTo(0.65, 2);
    }
    setShapeTextAutoFit(shape, 'normal');
    expect(getShapeBodyPrEffective(pres, shape).autoFitParams).toEqual({
      fontScale: 1,
      lnSpcReduction: 0,
    });
    setShapeTextAutoFit(shape, 'none');
    setShapeTextDirection(shape, 'horz');
    expect(getShapeBodyPrEffective(pres, shape)).toMatchObject({
      autoFit: 'none',
      autoFitParams: null,
      vert: 'horz',
    });
  });
});
