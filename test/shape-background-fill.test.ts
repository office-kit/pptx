import { describe, expect, it } from 'vitest';
import {
  createPresentation,
  addBlankSlide,
  addSlideShape,
  addSlideLine,
  setShapeSlideBackgroundFill,
  getShapeFill,
  getShapeFillEffective,
  getSlideXmlString,
  getSlides,
  getSlideShapes,
  savePresentation,
  loadPresentation,
  setShapeFill,
  setShapeNoFill,
  clearShapeFill,
  setShapeGradientFill,
  setShapePatternFill,
  setShapeImageFill,
  setShapeNoStroke,
  setShapeRotation,
  setShapeFlip,
  setSlideSize,
  inches,
  groupShapes,
  setShapeSize,
  setShapePosition,
} from '../src/api/index.ts';
import { SHAPE_ELEMENT, SLIDE_DOCUMENT } from '../src/api/_internal-symbols.ts';
import { parseXml } from '../src/internal/xml/index.ts';
import { renderSlideToRgba } from '../packages/preview/src/node.ts';

function deck() {
  const pres = createPresentation();
  setSlideSize(pres, { width: inches(10), height: inches(6) });
  const slide = addBlankSlide(pres);
  const shape = addSlideShape(slide, {
    preset: 'rect',
    x: inches(3),
    y: inches(2),
    w: inches(4),
    h: inches(2),
  });
  setShapeNoStroke(shape);
  return { pres, slide, shape };
}

describe('slide background fill', () => {
  it.each(['1', 'true', ' true '])('reads imported boolean %s', (value) => {
    const { pres, shape } = deck();
    setShapeSlideBackgroundFill(shape);
    shape[SHAPE_ELEMENT].attrs = shape[SHAPE_ELEMENT].attrs.map((a) =>
      a.name.localName === 'useBgFill' ? { ...a, value } : a,
    );
    expect(getShapeFill(shape)).toEqual({ kind: 'background' });
    expect(getShapeFillEffective(pres, shape)).toEqual({ kind: 'background' });
  });
  it('writes native useBgFill, preserves geometry and outline, and round trips', async () => {
    const { pres, slide, shape } = deck();
    const before = getSlideXmlString(slide);
    setShapeFill(shape, '#123456');
    setShapeSlideBackgroundFill(shape);
    expect(getShapeFill(shape)).toEqual({ kind: 'background' });
    expect(getShapeFillEffective(pres, shape)).toEqual({ kind: 'background' });
    expect(getSlideXmlString(slide).replace(' useBgFill="1"', '')).toBe(before);
    const loaded = await loadPresentation(await savePresentation(pres));
    expect(getShapeFill(getSlideShapes(getSlides(loaded)[0]!)[0]!)).toEqual({ kind: 'background' });
  });

  it('rejects connectors without mutating their XML', () => {
    const { slide } = deck();
    const line = addSlideLine(slide, {
      from: { x: inches(0), y: inches(0) },
      to: { x: inches(1), y: inches(1) },
    });
    const before = getSlideXmlString(slide);
    expect(() => setShapeSlideBackgroundFill(line)).toThrow('only ordinary shapes');
    expect(getSlideXmlString(slide)).toBe(before);
  });

  it('leaves background fill intact when a replacement fails validation', () => {
    const { slide, shape } = deck();
    setShapeSlideBackgroundFill(shape);
    const before = getSlideXmlString(slide);
    expect(() =>
      setShapeGradientFill(shape, { stops: [{ offset: 2, color: '#FF0000' }] }),
    ).toThrow();
    expect(getSlideXmlString(slide)).toBe(before);
  });

  it('all explicit fill setters and clear remove the background flag', () => {
    const { slide, shape } = deck();
    const updates = [
      () => setShapeFill(shape, '#FF0000'),
      () => setShapeNoFill(shape),
      () => clearShapeFill(shape),
      () => setShapePatternFill(shape, {}),
      () =>
        setShapeGradientFill(shape, {
          stops: [
            { offset: 0, color: '#FF0000' },
            { offset: 1, color: '#0000FF' },
          ],
        }),
      () =>
        setShapeImageFill(
          shape,
          new TextEncoder().encode(
            '<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"/>',
          ),
          { format: 'svg' },
        ),
    ];
    for (const update of updates) {
      setShapeSlideBackgroundFill(shape);
      update();
      expect(getShapeFill(shape).kind).not.toBe('background');
      expect(getSlideXmlString(slide)).not.toContain('useBgFill');
    }
  });

  it.each([0, 45, 90, 180])(
    'samples slide coordinates through rotation %s and flips, covering objects behind',
    (rotation) => {
      const { pres, slide, shape } = deck();
      const common = slide[SLIDE_DOCUMENT].root.children.find(
        (c) => c.kind === 'element' && c.name.localName === 'cSld',
      );
      if (!common || common.kind !== 'element') throw new Error('Missing cSld');
      common.children.unshift(
        parseXml(
          '<p:bg xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill></p:bgPr></p:bg>',
        ).root,
      );
      setShapeNoFill(shape);
      const reference = renderSlideToRgba(pres, slide, { width: 480 }).image;
      setShapeFill(shape, '#00FF00');
      const foreground = addSlideShape(slide, {
        preset: 'ellipse',
        x: inches(3),
        y: inches(2),
        w: inches(4),
        h: inches(2),
      });
      setShapeNoStroke(foreground);
      setShapeSlideBackgroundFill(foreground);
      setShapeRotation(foreground, rotation);
      setShapeFlip(foreground, { horizontal: true, vertical: true });
      if (rotation === 180) {
        const sibling = addSlideShape(slide, {
          preset: 'rect',
          x: inches(3),
          y: inches(2),
          w: inches(4),
          h: inches(2),
        });
        setShapeNoFill(sibling);
        setShapeNoStroke(sibling);
        const group = groupShapes([foreground, sibling]);
        setShapeSize(group, inches(5), inches(3));
        setShapePosition(group, inches(2.5), inches(1.5));
        setShapeRotation(group, 30);
        setShapeFlip(group, { horizontal: true });
      }
      const actual = renderSlideToRgba(pres, slide, { width: 480 }).image;
      for (const [x, y] of [
        [225, 135],
        [240, 144],
        [255, 153],
      ]) {
        const offset = (y! * actual.width + x!) * 4;
        expect(actual.data.slice(offset, offset + 4)).toEqual(
          reference.data.slice(offset, offset + 4),
        );
      }
      // A corner outside the foreground ellipse still exposes the green object.
      const corner = (98 * actual.width + 146) * 4;
      if (rotation !== 180)
        expect(Array.from(actual.data.slice(corner, corner + 4))).toEqual([0, 255, 0, 255]);
    },
  );
});
