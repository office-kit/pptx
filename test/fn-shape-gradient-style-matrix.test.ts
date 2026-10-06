import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  addSlideTextBox,
  getShapeGradientFillEffective,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url));

const style =
  '<p:style xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
  '<a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>' +
  '<a:fillRef idx="1"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:fillRef>' +
  '<a:effectRef idx="0"/><a:fontRef idx="minor"/>' +
  '</p:style>';

const matrixGradient =
  '<a:gradFill rotWithShape="0"><a:gsLst>' +
  '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="20000"/></a:schemeClr></a:gs>' +
  '<a:gs pos="100000"><a:srgbClr val="123456"><a:lumMod val="80000"/></a:srgbClr></a:gs>' +
  '</a:gsLst><a:lin ang="2700000" scaled="0"/></a:gradFill>';

const makeDeck = async (removeDirectFill = true) => {
  const pres = await loadPresentation(await readFile(fixturePath));
  const slide = getSlides(pres)[0]!;
  addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(2),
    h: inches(1),
    text: 'matrix gradient',
  });
  const source = unzipSync(await savePresentation(pres));
  const slideXml = strFromU8(source['ppt/slides/slide1.xml']!);
  const offset = slideXml.lastIndexOf('<p:txBody>');
  const withStyle = `${slideXml.slice(0, offset)}${style}${slideXml.slice(offset)}`;
  source['ppt/slides/slide1.xml'] = strToU8(
    removeDirectFill
      ? withStyle.replace(
          '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>',
          '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',
        )
      : withStyle,
  );
  const themeXml = strFromU8(source['ppt/theme/theme1.xml']!);
  source['ppt/theme/theme1.xml'] = strToU8(
    themeXml.replace(
      /<a:fillStyleLst>[\s\S]*?<\/a:fillStyleLst>/,
      `<a:fillStyleLst>${matrixGradient}${matrixGradient}${matrixGradient}</a:fillStyleLst>`,
    ),
  );
  return loadPresentation(zipSync(source));
};

describe('shape style-matrix gradient details', () => {
  it('resolves stops, transforms, angle, and survives save/load', async () => {
    const pres = await makeDeck();
    const shape = getSlideShapes(getSlides(pres)[0]!).at(-1)!;
    const gradient = getShapeGradientFillEffective(pres, shape);
    expect(gradient?.stops.map((stop) => stop.color)).toEqual(['scheme:accent1', '#123456']);
    expect(gradient?.stops[0]).toMatchObject({
      resolvedColor: expect.any(String),
      colorTransforms: [
        { kind: 'shade', value: 0.5 },
        { kind: 'tint', value: 0.2 },
      ],
    });
    expect(gradient?.stops[1]?.resolvedColor).toEqual('#0E2A45');
    expect(gradient?.stops[1]?.brightness).toBeCloseTo(-0.2, 10);
    expect(gradient?.angleDeg).toBe(45);
    expect(gradient?.rotateWithShape).toBe(false);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const restored = getSlideShapes(getSlides(reloaded)[0]!).at(-1)!;
    expect(getShapeGradientFillEffective(reloaded, restored)).toEqual(gradient);
  });

  it('keeps a direct noFill ahead of a style-matrix gradient', async () => {
    const pres = await makeDeck(false);
    const shape = getSlideShapes(getSlides(pres)[0]!).at(-1)!;
    expect(getShapeGradientFillEffective(pres, shape)).toBeNull();
  });
});
