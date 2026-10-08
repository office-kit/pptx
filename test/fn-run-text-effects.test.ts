// Character-level outline, shadow and glow — `<a:ln>` and `<a:effectLst>`
// inside a run's `<a:rPr>`, the decorative half of the text format.

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
  inches,
  loadPresentation,
  savePresentation,
  setShapeTextFormat,
  type TextFormat,
} from '../src/api/index.ts';

const textBox = (text = 'Styled 文字') => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const shape = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(4),
    h: inches(1),
    text,
  });
  return { pres, shape };
};

describe('character-level effects', () => {
  it('writes and reads an outline around the glyphs', () => {
    const { shape } = textBox();
    setShapeTextFormat(shape, { outline: { color: '#FF0000', widthEmu: 12700 } });

    expect(getShapeXmlString(shape)).toContain('<a:ln w="12700">');
    expect(getShapeRunFormat(shape, 0, 0)?.outline).toEqual({
      color: '#FF0000',
      widthEmu: 12700,
    });
  });

  it('writes and reads a shadow and a glow on the same run', () => {
    const { pres, shape } = textBox();
    setShapeTextFormat(shape, {
      shadow: { color: '#123456', blurEmu: 50800, offsetEmu: 38100, angleDeg: 90, opacity: 0.5 },
      glow: { color: '#00FF00', radiusEmu: 63500 },
    });

    const format = getShapeRunFormatEffective(pres, shape, 0, 0);
    expect(format.shadow).toEqual({
      alignment: 'tl',
      rotateWithShape: false,
      color: '#123456',
      blurEmu: 50800,
      offsetEmu: 38100,
      angleDeg: 90,
      opacity: 0.5,
    });
    expect(format.glow).toMatchObject({ color: '#00FF00', radiusEmu: 63500 });
    // CT_EffectList is a sequence: glow precedes outerShdw whichever order the
    // caller asked for them in.
    const xml = getShapeXmlString(shape);
    expect(xml.indexOf('<a:glow')).toBeLessThan(xml.indexOf('<a:outerShdw'));
  });

  it('round trips an authored shadow anchor and rotation independently of reflection', async () => {
    const { pres, shape } = textBox();
    setShapeTextFormat(shape, {
      shadow: { color: '#123456', alignment: 'ctr', rotateWithShape: true },
      reflection: {},
    });
    const restored = await loadPresentation(await savePresentation(pres));
    const format = getShapeRunFormat(getSlideShapes(getSlides(restored)[0]!)[0]!, 0, 0);
    expect(format?.shadow).toMatchObject({ alignment: 'ctr', rotateWithShape: true });
    expect(format?.reflection).toBeDefined();
    expect(getShapeXmlString(shape)).toContain('algn="ctr" rotWithShape="1"');
  });

  it('writes, reads, and removes a character reflection', async () => {
    const { pres, shape } = textBox();
    setShapeTextFormat(shape, {
      reflection: {
        blurEmu: 6350,
        offsetEmu: 0,
        angleDeg: 90,
        opacity: 0.003,
        startOpacity: 0.53,
        endPosition: 0.355,
        scaleY: -0.9,
        alignment: 'bl',
        rotateWithShape: false,
      },
    });

    const xml = getShapeXmlString(shape);
    expect(xml).toContain(
      '<a:reflection blurRad="6350" dist="0" dir="5400000" algn="bl" rotWithShape="0" stA="53000" endA="300" endPos="35500" sy="-90000"',
    );
    expect(getShapeRunFormatEffective(pres, shape, 0, 0).reflection).toEqual({
      blurEmu: 6350,
      offsetEmu: 0,
      angleDeg: 90,
      opacity: 0.003,
      startOpacity: 0.53,
      endPosition: 0.355,
      scaleY: -0.9,
      alignment: 'bl',
      rotateWithShape: false,
    });

    const reloaded = await loadPresentation(await savePresentation(pres));
    const restored = getSlideShapes(getSlides(reloaded)[0]!)[0]!;
    expect(getShapeRunFormat(restored, 0, 0)?.reflection).toEqual({
      blurEmu: 6350,
      offsetEmu: 0,
      angleDeg: 90,
      opacity: 0.003,
      startOpacity: 0.53,
      endPosition: 0.355,
      scaleY: -0.9,
      alignment: 'bl',
      rotateWithShape: false,
    });

    setShapeTextFormat(shape, { shadow: { color: '#000000' } });
    setShapeTextFormat(shape, { reflection: null });
    expect(getShapeXmlString(shape)).not.toContain('<a:reflection');
    expect(getShapeRunFormat(shape, 0, 0)?.shadow).toMatchObject({ color: '#000000' });
  });

  it('round trips every reflection transform and preserves schema defaults', async () => {
    const { pres, shape } = textBox();
    const reflection = {
      blurEmu: 6350,
      offsetEmu: 12700,
      angleDeg: 30,
      startOpacity: 0.7,
      startPosition: 0.1,
      opacity: 0.05,
      endPosition: 0.8,
      fadeDirection: 120,
      scaleX: 1.5,
      scaleY: -0.9,
      skewX: 12,
      skewY: -8,
      alignment: 'tr' as const,
      rotateWithShape: false,
    };
    setShapeTextFormat(shape, { reflection });
    const restored = await loadPresentation(await savePresentation(pres));
    expect(
      getShapeRunFormat(getSlideShapes(getSlides(restored)[0]!)[0]!, 0, 0)?.reflection,
    ).toEqual(reflection);
    setShapeTextFormat(shape, { reflection: {} });
    expect(getShapeRunFormat(shape, 0, 0)?.reflection).toMatchObject({
      angleDeg: 0,
      alignment: 'b',
      rotateWithShape: true,
    });
  });

  it('validates fixed angles and percentages against the DrawingML bounds', () => {
    const { shape } = textBox();
    for (const skewX of [-90, 90, 100, Number.NaN]) {
      expect(() => setShapeTextFormat(shape, { reflection: { skewX } })).toThrow();
    }
    expect(() => setShapeTextFormat(shape, { reflection: { scaleX: 1e20 } })).toThrow();
    setShapeTextFormat(shape, { reflection: { angleDeg: 359.99999999, fadeDirection: -90 } });
    expect(getShapeRunFormat(shape, 0, 0)?.reflection).toMatchObject({
      angleDeg: 0,
      fadeDirection: 270,
    });
  });

  it('removes reflection independently and with a whole-format reset', () => {
    const { shape } = textBox();
    setShapeTextFormat(shape, {
      reflection: {},
      shadow: { color: '#000000' },
      glow: { color: '#FFFFFF' },
    });
    setShapeTextFormat(shape, { reflection: null });
    expect(getShapeRunFormat(shape, 0, 0)).toMatchObject({
      shadow: { color: '#000000' },
      glow: { color: '#FFFFFF' },
    });
    expect(getShapeRunFormat(shape, 0, 0)?.reflection).toBeUndefined();
    setShapeTextFormat(shape, { reflection: {} });
    setShapeTextFormat(shape, { size: 18 }, { reset: true });
    expect(getShapeRunFormat(shape, 0, 0)?.reflection).toBeUndefined();
  });

  it('accepts reflection magnification and rejects non-finite geometry', () => {
    const { shape } = textBox();
    expect(() => setShapeTextFormat(shape, { reflection: { scaleY: 1.5 } })).not.toThrow();
    expect(() => setShapeTextFormat(shape, { reflection: { angleDeg: Number.NaN } })).toThrow(
      /angleDeg must be finite/,
    );
    expect(() =>
      setShapeTextFormat(shape, { reflection: { scaleY: Number.POSITIVE_INFINITY } }),
    ).toThrow(/scaleY must be finite/);
  });

  it('keeps the run’s own child order — ln, effectLst, then the fonts', () => {
    const { shape } = textBox();
    setShapeTextFormat(shape, {
      font: 'Arial',
      color: '#000080',
      glow: { color: '#FFFF00' },
      outline: { color: '#FFFFFF', widthEmu: 9525 },
    });

    const xml = getShapeXmlString(shape);
    // CT_TextCharacterProperties: ln → fill → effectLst → … → latin.
    expect(xml.indexOf('<a:ln ')).toBeLessThan(xml.indexOf('<a:solidFill>'));
    expect(xml.indexOf('<a:solidFill>')).toBeLessThan(xml.indexOf('<a:effectLst>'));
    expect(xml.indexOf('<a:effectLst>')).toBeLessThan(xml.indexOf('<a:latin'));
  });

  it('removes one effect without disturbing the other', () => {
    const { pres, shape } = textBox();
    setShapeTextFormat(shape, {
      shadow: { color: '#000000' },
      glow: { color: '#00FF00' },
      outline: { color: '#FF0000' },
    });

    setShapeTextFormat(shape, { glow: null });

    const format = getShapeRunFormatEffective(pres, shape, 0, 0);
    expect(format.glow).toBeUndefined();
    expect(format.shadow).toMatchObject({ color: '#000000' });
    expect(format.outline).toMatchObject({ color: '#FF0000' });
  });

  it('drops the effect list with its last effect, so the run inherits again', () => {
    const { shape } = textBox();
    setShapeTextFormat(shape, { glow: { color: '#00FF00' } });
    setShapeTextFormat(shape, { glow: null });

    // An empty `<a:effectLst>` states "no effects", which is not the same as
    // saying nothing at all.
    expect(getShapeXmlString(shape)).not.toContain('effectLst');
  });

  it('removes the outline on null', () => {
    const { shape } = textBox();
    setShapeTextFormat(shape, { outline: { color: '#FF0000', widthEmu: 12700 } });
    setShapeTextFormat(shape, { outline: null });

    expect(getShapeXmlString(shape)).not.toContain('<a:ln');
    expect(getShapeRunFormat(shape, 0, 0)?.outline).toBeUndefined();
  });

  it('carries a glow’s opacity, which the reader reports', () => {
    const { pres, shape } = textBox();
    setShapeTextFormat(shape, { glow: { color: '#00FF00', radiusEmu: 63500, opacity: 0.4 } });

    expect(getShapeRunFormatEffective(pres, shape, 0, 0).glow).toEqual({
      color: '#00FF00',
      radiusEmu: 63500,
      opacity: 0.4,
    });
  });

  it('survives the save/load round trip', async () => {
    const { pres, shape } = textBox();
    const format: TextFormat = {
      outline: { color: '#FF0000', widthEmu: 12700 },
      shadow: { color: '#123456', blurEmu: 50800, offsetEmu: 38100, angleDeg: 90 },
      glow: { color: '#00FF00', radiusEmu: 63500 },
    };
    setShapeTextFormat(shape, format);
    const before = getShapeRunFormat(shape, 0, 0);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const saved = getSlideShapes(getSlides(reloaded)[0]!)[0]!;

    expect(getShapeRunFormat(saved, 0, 0)).toEqual(before);
  });

  it('is cleared with the rest of the run’s look by reset', () => {
    const { shape } = textBox();
    setShapeTextFormat(shape, {
      outline: { color: '#FF0000' },
      glow: { color: '#00FF00' },
      shadow: { color: '#000000' },
    });

    setShapeTextFormat(shape, { size: 18 }, { reset: true });

    const xml = getShapeXmlString(shape);
    expect(xml).not.toContain('<a:ln');
    expect(xml).not.toContain('effectLst');
  });
});
