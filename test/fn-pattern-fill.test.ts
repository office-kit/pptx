// Preset pattern fill on a shape.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import {
  addSlideShape,
  getShapeFill,
  getSlideShapes,
  getSlideXmlString,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapePatternFill,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const slideXml = async (bytes: Uint8Array, slideIndex: number): Promise<string> => {
  const pres = await loadPresentation(bytes);
  return getSlideXmlString(getSlides(pres)[slideIndex]!);
};

describe('fn API: setShapePatternFill', () => {
  it('writes <a:pattFill> with the preset + fg/bg colors', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(2),
    });
    setShapePatternFill(shape, {
      preset: 'pct50',
      foreground: '#FF0000',
      background: '#FFFFFF',
    });
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('<a:pattFill ');
    expect(xml).toContain('prst="pct50"');
    expect(xml).toContain('FF0000');
    expect(xml).toContain('FFFFFF');
  });

  it('getShapeFill reports pattern after setShapePatternFill', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(2),
    });
    setShapePatternFill(shape, {
      preset: 'dkUpDiag',
      foreground: '#000000',
      background: '#FFFFFF',
    });
    expect(getShapeFill(shape).kind).toBe('pattern');
  });

  it('replaces any previous fill choice', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(2),
    });
    setShapePatternFill(shape, {
      preset: 'pct25',
      foreground: '#FF0000',
      background: '#FFFFFF',
    });
    setShapePatternFill(shape, {
      preset: 'pct75',
      foreground: '#0000FF',
      background: '#FFFFFF',
    });
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('pct75');
    expect(xml).not.toContain('pct25');
  });
  it('preserves theme references and transforms when only the preset or one color changes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(2),
    });
    setShapePatternFill(shape, { preset: 'pct5', foreground: 'accent1', background: 'bg1' });
    const parts = unzipSync(await savePresentation(pres));
    const path = 'ppt/slides/slide1.xml';
    parts[path] = strToU8(
      strFromU8(parts[path]!).replace(
        '<a:schemeClr val="accent1"/>',
        '<a:schemeClr val="accent1"><a:lumMod val="75000"/><a:alpha val="80000"/></a:schemeClr>',
      ),
    );
    const imported = await loadPresentation(zipSync(parts));
    const importedSlide = getSlides(imported)[0]!;
    const target = getSlideShapes(importedSlide).at(-1)!;
    const original = getSlideXmlString(importedSlide);
    setShapePatternFill(target, { preset: 'wave' });
    expect(getSlideXmlString(importedSlide)).toBe(original.replace('prst="pct5"', 'prst="wave"'));
    setShapePatternFill(target, { background: '#123456' });
    const expected = original
      .replace('prst="pct5"', 'prst="wave"')
      .replace('<a:schemeClr val="bg1"/>', '<a:srgbClr val="123456"/>');
    expect(getSlideXmlString(importedSlide)).toBe(expected);
    expect(await slideXml(await savePresentation(imported), 0)).toBe(expected);
  });

  it('rejects invalid pattern edits before changing the fill', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(2),
    });
    setShapePatternFill(shape, { preset: 'pct5', foreground: 'accent1', background: 'bg1' });
    const before = getSlideXmlString(slide);
    expect(() =>
      setShapePatternFill(shape, {
        // @ts-expect-error untrusted runtime input
        preset: 'invalid',
        foreground: '#000000',
        background: '#FFFFFF',
      }),
    ).toThrow();
    expect(getSlideXmlString(slide)).toBe(before);
    expect(await slideXml(await savePresentation(pres), 0)).toBe(before);
  });

  it('uses native defaults when creating a pattern with unspecified settings', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(2),
    });
    setShapePatternFill(shape, {});
    expect(getSlideXmlString(slide)).toContain(
      '<a:pattFill prst="pct5"><a:fgClr><a:schemeClr val="accent1"/></a:fgClr><a:bgClr><a:schemeClr val="bg1"/></a:bgClr></a:pattFill>',
    );
  });
});
