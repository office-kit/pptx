import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { resolveDrawingColor, resolveDrawingColorOpacity } from '../src/api/index.ts';
import { parseEffectList, parseRPrLikeElement } from '../src/api/fn/shape-color.ts';
import { NS, firstChildElement, parseXml, qname } from '../src/internal/xml/index.ts';

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

  it('resolves non-sRGB effect colors in the effect parser', () => {
    const effects = parseXml(
      `<a:effectLst xmlns:a="${A}"><a:glow rad="100"><a:scrgbClr r="100%" g="0%" b="0%"/></a:glow><a:outerShdw blurRad="0" dist="0" dir="0"><a:hslClr hue="7200000" sat="100%" lum="50%"/></a:outerShdw></a:effectLst>`,
    ).root;
    expect(parseEffectList(effects, null)).toEqual([
      { kind: 'glow', color: '#FF0000', radiusEmu: 100 },
      {
        kind: 'outerShdw',
        color: '#00FF00',
        blurEmu: 0,
        distEmu: 0,
        angleDeg: 0,
      },
    ]);
  });

  it('reads percent-suffixed reflection opacity', () => {
    const effects = parseXml(
      `<a:effectLst xmlns:a="${A}"><a:reflection blurRad="0" dist="0" dir="0" endA="50%"/></a:effectLst>`,
    ).root;
    expect(parseEffectList(effects, null)[0]).toMatchObject({ opacity: 0.5 });
  });

  it('reads native reflection attributes without dropping the extra geometry', () => {
    const effects = parseXml(
      `<a:effectLst xmlns:a="${A}"><a:reflection blurRad="6350" dist="0" dir="5400000" stA="53000" endA="300" endPos="35500" sy="-90000" algn="bl" rotWithShape="0"/></a:effectLst>`,
    ).root;
    expect(parseEffectList(effects, null)).toEqual([
      {
        kind: 'reflection',
        blurEmu: 6350,
        distEmu: 0,
        angleDeg: 90,
        opacity: 0.003,
        startOpacity: 0.53,
        endPosition: 0.355,
        scaleY: -0.9,
        alignment: 'bl',
        rotateWithShape: false,
      },
    ]);
  });

  it('reads the captured native WordArt reflection fixture', async () => {
    const xml = parseXml(
      await readFile(
        new URL('./fixtures/native/wordart-accent5-gradient-reflection-shape.xml', import.meta.url),
        'utf8',
      ),
    ).root;
    const txBody = firstChildElement(xml, qname('p', 'txBody', NS.pml));
    const paragraph = txBody && firstChildElement(txBody, qname('a', 'p', NS.dml));
    const run = paragraph && firstChildElement(paragraph, qname('a', 'r', NS.dml));
    const rPr = run && firstChildElement(run, qname('a', 'rPr', NS.dml));
    expect(rPr).not.toBeNull();
    expect(parseRPrLikeElement(rPr!).reflection).toEqual({
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
  });
});
