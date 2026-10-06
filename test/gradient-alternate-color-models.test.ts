import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideShape,
  createPresentation,
  getShapeFillColorResolved,
  getShapeFill,
  getShapeGradientFill,
  getShapeGradientFillEffective,
  getShapeStroke,
  getShapeStrokeColorResolved,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeFill,
  setShapeGradientFill,
  setShapeStroke,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';

const replaceOnce = (source: string, target: string, replacement: string): string => {
  const index = source.indexOf(target);
  if (index < 0) throw new Error(`XML fragment not found: ${target}`);
  return source.slice(0, index) + replacement + source.slice(index + target.length);
};

describe('imported alternate DrawingML color models', () => {
  it('reads scrgbClr and hslClr in solid fill, stroke, and gradient stops', async () => {
    const presentation = createPresentation();
    const slide = addBlankSlide(presentation);
    const solid = addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
    });
    setShapeFill(solid, '#FF0000');
    setShapeStroke(solid, { color: '#FF0000', widthEmu: 12700 });
    const gradient = addSlideShape(slide, {
      preset: 'rect',
      x: inches(4),
      y: inches(1),
      w: inches(2),
      h: inches(1),
    });
    setShapeGradientFill(gradient, {
      stops: [
        { offset: 0, color: '#FF0000' },
        { offset: 1, color: '#00FF00' },
      ],
    });

    const zip = readZip(await savePresentation(presentation));
    const slideEntry = zip.entries.find((entry) => entry.name === 'ppt/slides/slide1.xml');
    if (!slideEntry) throw new Error('slide XML missing');
    let xml = new TextDecoder().decode(slideEntry.data);
    xml = replaceOnce(
      xml,
      '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>',
      '<a:solidFill><a:scrgbClr r="100%" g="0%" b="0%"/></a:solidFill>',
    );
    xml = replaceOnce(
      xml,
      '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>',
      '<a:solidFill><a:hslClr hue="0" sat="100%" lum="50%"/></a:solidFill>',
    );
    xml = replaceOnce(
      xml,
      '<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>',
      '<a:gs pos="0"><a:scrgbClr r="100%" g="0%" b="0%"><a:shade val="50000"/></a:scrgbClr></a:gs>',
    );
    xml = replaceOnce(
      xml,
      '<a:gs pos="100000"><a:srgbClr val="00FF00"/></a:gs>',
      '<a:gs pos="100000"><a:hslClr hue="7200000" sat="100%" lum="50%"/></a:gs>',
    );
    const loaded = await loadPresentation(
      writeZip(
        zip.entries.map((entry) =>
          entry.name === slideEntry.name
            ? { ...entry, data: new TextEncoder().encode(xml) }
            : entry,
        ),
      ),
    );
    const assertColors = (reopened: Awaited<ReturnType<typeof loadPresentation>>) => {
      const shapes = getSlideShapes(getSlides(reopened)[0]!);
      expect(getShapeFill(shapes[0]!)).toEqual({ kind: 'solid', color: '#FF0000' });
      expect(getShapeStroke(shapes[0]!)).toMatchObject({ kind: 'solid', color: '#FF0000' });
      expect(getShapeFillColorResolved(reopened, shapes[0]!)).toBe('#FF0000');
      expect(getShapeStrokeColorResolved(reopened, shapes[0]!)).toBe('#FF0000');
      expect(getShapeGradientFill(shapes[1]!)?.stops).toEqual([
        {
          offset: 0,
          color: '#FF0000',
          colorTransforms: [{ kind: 'shade', value: 0.5 }],
        },
        { offset: 1, color: '#00FF00' },
      ]);
      expect(getShapeGradientFillEffective(reopened, shapes[1]!)?.stops).toMatchObject([
        { color: '#FF0000', resolvedColor: '#BC0000' },
        { color: '#00FF00', resolvedColor: '#00FF00' },
      ]);
    };
    assertColors(loaded);
    assertColors(await loadPresentation(await savePresentation(loaded)));
  });
});
