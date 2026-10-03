import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  addBlankSlide,
  addSlideShape,
  createPresentation,
  getShapeFill,
  getShapeFillEffective,
  getShapeStroke,
  getShapeStrokeEffective,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeFill,
  setShapeStroke,
} from '../src/api/index.ts';

describe.each(['slide', 'layout', 'master'])('%s paint', (source) => {
  it.each([
    ['<a:prstClr val="red"><a:shade val="50000"/></a:prstClr>', '#FF0000'],
    ['<a:sysClr val="windowText" lastClr="123456"/>', '#123456'],
  ])('reads the base fill and outline color from %s', async (color, expected) => {
    const pres = createPresentation();
    const shape = addSlideShape(addBlankSlide(pres), {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
    });
    setShapeFill(shape, '#112233');
    setShapeStroke(shape, { color: '#112233', widthEmu: 25400 });
    const zip = unzipSync(await savePresentation(pres));
    zip['ppt/slides/slide1.xml'] = strToU8(
      strFromU8(zip['ppt/slides/slide1.xml']!).replaceAll('<a:srgbClr val="112233"/>', color),
    );
    if (source !== 'slide') {
      const slideXml = strFromU8(zip['ppt/slides/slide1.xml']!);
      const shapeXml = slideXml.match(/<p:sp>.*?<\/p:sp>/)?.[0];
      if (!shapeXml) throw new Error('authored shape missing');
      const placeholder = shapeXml.replace('<p:nvPr/>', '<p:nvPr><p:ph idx="42"/></p:nvPr>');
      expect(placeholder).toContain('<p:ph idx="42"/>');
      const sourceName =
        source === 'layout'
          ? 'ppt/slideLayouts/slideLayout1.xml'
          : 'ppt/slideMasters/slideMaster1.xml';
      zip[sourceName] = strToU8(
        strFromU8(zip[sourceName]!).replace('</p:spTree>', `${placeholder}</p:spTree>`),
      );
      const inheriting = placeholder
        .replace(/<a:solidFill>.*?<\/a:solidFill>/g, '')
        .replace(/<a:ln\b[^>]*>.*?<\/a:ln>/g, '');
      zip['ppt/slides/slide1.xml'] = strToU8(slideXml.replace(shapeXml, inheriting));
    }
    const imported = await loadPresentation(zipSync(zip));
    const assertColors = (deck: typeof pres) => {
      const restored = getSlideShapes(getSlides(deck)[0]!)[0]!;
      if (source !== 'slide') {
        expect(getShapeFill(restored).kind).toBe('inherit');
        expect(getShapeStroke(restored).kind).toBe('inherit');
      }
      expect(getShapeFillEffective(deck, restored)).toEqual({ kind: 'solid', color: expected });
      expect(getShapeStrokeEffective(deck, restored)).toEqual({
        kind: 'solid',
        color: expected,
        widthEmu: 25400,
      });
    };
    assertColors(imported);
    assertColors(await loadPresentation(await savePresentation(imported)));
  });
});
