import { describe, expect, it } from 'vitest';
import { resolveDrawingColor, resolveDrawingColorOpacity } from '../src/api/index.ts';
import { parseEffectList } from '../src/api/fn/shape-color.ts';
import { parseXml } from '../src/internal/xml/index.ts';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const parseColor = (transform: string) =>
  parseXml(`<a:srgbClr xmlns:a="${A}" val="FF0000">${transform}</a:srgbClr>`).root;

describe('DrawingML color transform percentage lexical forms', () => {
  it('reads percent-suffixed color transforms as percentages', () => {
    const color = parseColor('<a:shade val="50%"/>');
    expect(resolveDrawingColor(color, null)).toBe('#BC0000');
  });

  it('reads the smallest fixed-point integer without treating it as one', () => {
    const color = parseColor('<a:lumMod val="1"/>');
    expect(resolveDrawingColor(color, null)).toBe('#000000');
  });

  it('keeps fixed-point semantics for decimal spellings above one', () => {
    const color = parseColor('<a:lumMod val="50000.0"/>');
    expect(resolveDrawingColor(color, null)).toBe('#800000');
  });

  it('reads hueOff as an ST_Angle rather than a percentage', () => {
    const color = parseColor('<a:hueOff val="9000000"/>');
    expect(resolveDrawingColor(color, null)).toBe('#00FF80');
  });

  it('accepts the smallest signed ST_Angle value for hueOff', () => {
    const color = parseColor('<a:hueOff val="1"/>');
    expect(resolveDrawingColor(color, null)).toBe('#FF0000');
  });

  it('reads percent-suffixed alpha transforms as opacity', () => {
    const color = parseColor('<a:alpha val="50%"/>');
    expect(resolveDrawingColorOpacity(color)).toBeCloseTo(0.5, 6);
  });

  it('reads percent-suffixed effect color opacity', () => {
    const effects = parseXml(
      `<a:effectLst xmlns:a="${A}"><a:glow rad="100"><a:srgbClr val="FF0000"><a:alpha val="50%"/></a:srgbClr></a:glow></a:effectLst>`,
    ).root;
    expect(parseEffectList(effects, null)).toEqual([
      { kind: 'glow', color: '#FF0000', radiusEmu: 100, opacity: 0.5 },
    ]);
  });

  it('reads percent-suffixed reflection opacity', () => {
    const effects = parseXml(
      `<a:effectLst xmlns:a="${A}"><a:reflection blurRad="0" dist="0" dir="0" endA="50%"/></a:effectLst>`,
    ).root;
    expect(parseEffectList(effects, null)[0]).toMatchObject({ opacity: 0.5 });
  });
});
