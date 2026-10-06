// Color transforms on a run's solid fill, outline and effect colors — the
// theme tints PowerPoint writes as `<a:schemeClr val="accent2"><a:lumMod/>…`.

import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  getShapeRunFormat,
  getShapeRunFormatEffective,
  getShapeXmlString,
  getSlides,
  getSlideShapes,
  getSlideXmlString,
  inches,
  loadPresentation,
  savePresentation,
  setShapeShadow,
  setShapeTextFormat,
  toWritableTextFormat,
  type ColorTransform,
  type TextFormat,
} from '../src/api/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const textBox = () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const shape = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(4),
    h: inches(1),
    text: 'Outline title',
  });
  return { pres, slide, shape };
};

const lighter: ColorTransform[] = [
  { kind: 'lumMod', value: 0.4 },
  { kind: 'lumOff', value: 0.6 },
];
const darker: ColorTransform[] = [{ kind: 'lumMod', value: 0.5 }];

const FORMAT: TextFormat = {
  color: 'accent2',
  colorTransforms: lighter,
  outline: { color: 'accent3', colorTransforms: darker, widthEmu: 12700 },
  shadow: { color: 'dk1', colorTransforms: darker, opacity: 0.4, blurEmu: 38100 },
  innerShadow: { color: 'accent3', colorTransforms: darker, blurEmu: 177800, offsetEmu: 0 },
  glow: { color: 'accent1', colorTransforms: [{ kind: 'tint', value: 0.01 }], opacity: 0.4 },
};

const compact = (xml: string) => xml.replace(/>\s+</g, '><').replace(/\s+\/>/g, '/>');

describe('run color transforms', () => {
  it('writes the transforms as children of each color, alpha last', () => {
    const { shape } = textBox();
    setShapeTextFormat(shape, FORMAT);
    const xml = compact(getShapeXmlString(shape));
    expect(xml).toContain(
      '<a:solidFill><a:schemeClr val="accent2"><a:lumMod val="40000"/><a:lumOff val="60000"/></a:schemeClr></a:solidFill>',
    );
    expect(xml).toContain(
      '<a:ln w="12700"><a:solidFill><a:schemeClr val="accent3"><a:lumMod val="50000"/></a:schemeClr></a:solidFill></a:ln>',
    );
    // PowerPoint's order: the tint, then the shadow's opacity.
    expect(xml).toContain(
      '<a:schemeClr val="dk1"><a:lumMod val="50000"/><a:alpha val="40000"/></a:schemeClr>',
    );
    expect(xml).toContain(
      '<a:innerShdw blurRad="177800" dist="0" dir="2700000"><a:schemeClr val="accent3"><a:lumMod val="50000"/></a:schemeClr></a:innerShdw>',
    );
    expect(xml).toContain(
      '<a:glow rad="63500"><a:schemeClr val="accent1"><a:tint val="1000"/><a:alpha val="40000"/></a:schemeClr></a:glow>',
    );
  });

  it('round trips through save and load, and through toWritableTextFormat', async () => {
    const { pres, shape } = textBox();
    setShapeTextFormat(shape, FORMAT);
    const reloaded = await loadPresentation(await savePresentation(pres));
    const restored = getSlideShapes(getSlides(reloaded)[0]!)[0]!;
    const read = getShapeRunFormat(restored, 0, 0)!;
    expect(read.color).toBe('accent2');
    expect(read.colorTransforms).toEqual(lighter);
    expect(read.outline).toEqual(FORMAT.outline);
    // The shadow reports its alpha both ways, as gradient stops do.
    expect(read.shadow).toMatchObject({
      color: 'dk1',
      colorTransforms: [...darker, { kind: 'alpha', value: 0.4 }],
      opacity: 0.4,
    });
    expect(read.innerShadow).toMatchObject({ color: 'accent3', colorTransforms: darker });
    expect(read.glow).toMatchObject({ color: 'accent1', opacity: 0.4 });

    const before = compact(getShapeXmlString(restored));
    setShapeTextFormat(restored, toWritableTextFormat(read));
    expect(compact(getShapeXmlString(restored))).toBe(before);
  });

  it('resolves the transforms instead of reporting them once a theme applies', () => {
    const { pres, shape } = textBox();
    setShapeTextFormat(shape, { color: 'accent1', colorTransforms: darker });
    const effective = getShapeRunFormatEffective(pres, shape, 0, 0);
    expect(effective.color).toMatch(/^#[0-9A-F]{6}$/);
    expect(effective.colorTransforms).toBeUndefined();
  });

  it('reports an effect color without a theme as its token, not an empty string', () => {
    const { shape } = textBox();
    setShapeTextFormat(shape, { shadow: { color: 'accent5' } });
    expect(getShapeRunFormat(shape, 0, 0)?.shadow?.color).toBe('accent5');
  });

  it('replaces the transforms a previous color carried', () => {
    const { shape } = textBox();
    setShapeTextFormat(shape, FORMAT);
    setShapeTextFormat(shape, { color: 'accent2' });
    expect(getShapeRunFormat(shape, 0, 0)?.colorTransforms).toBeUndefined();
    setShapeTextFormat(shape, { outline: { color: 'accent3', colorTransforms: lighter } });
    expect(getShapeRunFormat(shape, 0, 0)?.outline?.colorTransforms).toEqual(lighter);
  });

  it('lets an explicit opacity override the alpha the transforms state', () => {
    const { shape } = textBox();
    setShapeTextFormat(shape, {
      shadow: {
        color: 'dk1',
        colorTransforms: [...darker, { kind: 'alpha', value: 0.4 }],
        opacity: 0.7,
      },
    });
    expect(compact(getShapeXmlString(shape))).toContain(
      '<a:schemeClr val="dk1"><a:lumMod val="50000"/><a:alpha val="70000"/></a:schemeClr>',
    );
  });

  it('applies to shape-level effects through the same options', () => {
    const { shape } = textBox();
    setShapeShadow(shape, { color: 'accent1', colorTransforms: darker });
    expect(compact(getShapeXmlString(shape))).toContain(
      '<a:schemeClr val="accent1"><a:lumMod val="50000"/></a:schemeClr>',
    );
  });

  it('rejects transforms without a color, and invalid values, before editing', () => {
    const { shape } = textBox();
    const before = getShapeXmlString(shape);
    expect(() => setShapeTextFormat(shape, { colorTransforms: darker })).toThrow(
      /colorTransforms requires color/,
    );
    expect(() => setShapeTextFormat(shape, { color: null, colorTransforms: darker })).toThrow(
      /colorTransforms requires color/,
    );
    expect(() =>
      setShapeTextFormat(shape, { bold: true, outline: { colorTransforms: darker } }),
    ).toThrow(/outline.colorTransforms requires outline.color/);
    expect(() =>
      setShapeTextFormat(shape, {
        bold: true,
        glow: { color: 'accent1', colorTransforms: [{ kind: 'tint', value: 2 }] },
      }),
    ).toThrow(/invalid tint/);
    expect(() => setShapeShadow(shape, { colorTransforms: darker })).toThrow(
      /colorTransforms requires color/,
    );
    expect(getShapeXmlString(shape)).toBe(before);
  });

  it.skipIf(!isSchemaValidationAvailable())('validates against pml.xsd', () => {
    const { slide, shape } = textBox();
    setShapeTextFormat(shape, FORMAT);
    expectSchemaValid(getSlideXmlString(slide), 'pml');
  });
});
