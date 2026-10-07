// Extended run-format attributes: strike, spc (character spacing),
// kern (kerning threshold), baseline (super / sub), cap, highlight.
// These are part of ECMA-376's CT_TextCharacterProperties surface
// (§17.18.83) and round-trip through `setShapeRunFormat` /
// `getShapeRunFormat`.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  getShapeRunFormat,
  getShapeXmlString,
  getSlides,
  inches,
  loadPresentation,
  setShapeRunFormat,
  savePresentation,
  getSlideShapes,
} from '../src/api/index.ts';
import { parseRPrLikeElement } from '../src/api/fn/shape-color.ts';
import { parseXml } from '../src/internal/xml/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: extended run-format properties', () => {
  it.each([
    ['1', 0.00001],
    ['-1', -0.00001],
    ['1%', 0.01],
    ['-25%', -0.25],
    ['30000', 0.3],
  ])('reads baseline="%s" as a DrawingML percentage', (value, expected) => {
    const xml = parseXml(
      `<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" baseline="${value}"/>`,
    );
    expect(parseRPrLikeElement(xml.root).baseline).toBeCloseTo(expected, 8);
  });

  it('retains a small baseline offset through save and reload', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const shape = addSlideTextBox(getSlides(pres)[0]!, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      text: 'offset',
    });
    setShapeRunFormat(shape, 0, 0, { baseline: 0.00001 });
    expect(getShapeRunFormat(shape, 0, 0)!.baseline).toBeCloseTo(0.00001, 8);
    const loaded = await loadPresentation(await savePresentation(pres));
    expect(
      getShapeRunFormat(getSlideShapes(getSlides(loaded)[0]!).at(-1)!, 0, 0)!.baseline,
    ).toBeCloseTo(0.00001, 8);
  });

  it('null removes bold and spacing, where false and 0 write them off', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const shape = addSlideTextBox(getSlides(pres)[0]!, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      text: 'spaced',
    });
    setShapeRunFormat(shape, 0, 0, { bold: false, spc: 0 });
    expect(getShapeXmlString(shape)).toMatch(/<a:rPr[^>]* b="0"[^>]* spc="0"/);
    setShapeRunFormat(shape, 0, 0, { bold: null, spc: null });
    expect(getShapeXmlString(shape)).not.toMatch(/ (b|spc)="/);
    const format = getShapeRunFormat(shape, 0, 0) ?? {};
    expect('bold' in format || 'spc' in format).toBe(false);
  });

  it.each([
    ['1', true],
    ['true', true],
    ['0', false],
    ['false', false],
  ])('reads bold and italic XML boolean "%s"', (value, enabled) => {
    const xml = parseXml(
      `<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" b="${value}" i="${value}"/>`,
    );
    expect(parseRPrLikeElement(xml.root)).toMatchObject({ bold: enabled, italic: enabled });
  });
  it.each([true, false])(
    'preserves equalized character height (%s) through save and reload',
    async (enabled) => {
      const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
      const slide = getSlides(pres)[0]!;
      const tb = addSlideTextBox(slide, {
        x: inches(0),
        y: inches(0),
        w: inches(4),
        h: inches(2),
        text: 'AaBb',
      });
      setShapeRunFormat(tb, 0, 0, { normalizeHeight: enabled });
      setShapeRunFormat(tb, 0, 0, { bold: true });
      const loaded = await loadPresentation(await savePresentation(pres));
      const shape = getSlideShapes(getSlides(loaded)[0]!).at(-1)!;
      expect(getShapeRunFormat(shape, 0, 0)).toMatchObject({
        normalizeHeight: enabled,
        bold: true,
      });
    },
  );
  it.each([
    ['1', true],
    ['true', true],
    ['0', false],
    ['false', false],
  ])('reads normalizeH="%s"', (value, enabled) => {
    const xml = parseXml(
      `<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" normalizeH="${value}"/>`,
    );
    expect(parseRPrLikeElement(xml.root).normalizeHeight).toBe(enabled);
  });
  it('round-trips strike, spc, kern, baseline, cap, highlight', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      text: 'extended',
    });
    setShapeRunFormat(tb, 0, 0, {
      strike: true,
      spc: 200,
      kern: 1200,
      baseline: 0.3,
      cap: 'all',
      highlight: '#FFFF00',
    });
    const fmt = getShapeRunFormat(tb, 0, 0);
    expect(fmt).not.toBeNull();
    expect(fmt!.strike).toBe(true);
    expect(fmt!.spc).toBe(200);
    expect(fmt!.kern).toBe(1200);
    expect(fmt!.baseline).toBeCloseTo(0.3, 4);
    expect(fmt!.cap).toBe('all');
    expect(fmt!.highlight).toBe('#FFFF00');
  });

  it('strike accepts both boolean shorthand and explicit dblStrike', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 's',
    });
    setShapeRunFormat(tb, 0, 0, { strike: true });
    expect(getShapeRunFormat(tb, 0, 0)!.strike).toBe(true);
    setShapeRunFormat(tb, 0, 0, { strike: 'dblStrike' });
    expect(getShapeRunFormat(tb, 0, 0)!.strike).toBe('dblStrike');
    setShapeRunFormat(tb, 0, 0, { strike: false });
    expect(getShapeRunFormat(tb, 0, 0)!.strike).toBe(false);
  });
});
