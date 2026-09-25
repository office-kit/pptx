// Character-level outline, shadow and glow — `<a:ln>` and `<a:effectLst>`
// inside a run's `<a:rPr>`, the WordArt half of the text format.

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

const textBox = (text = 'WordArt 文字') => {
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
