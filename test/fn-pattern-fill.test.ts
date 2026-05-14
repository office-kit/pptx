// Preset pattern fill on a shape.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  Presentation,
  _internalPackageOf,
  addSlideShape,
  getShapeFill,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapePatternFill,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const slideXml = async (bytes: Uint8Array, slideIndex: number): Promise<string> => {
  const pres = await Presentation.load(bytes);
  const pkg = _internalPackageOf(pres);
  const part = pkg.parts.find((p) => p.name === `/ppt/slides/slide${slideIndex + 1}.xml`);
  if (!part) throw new Error(`slide${slideIndex + 1}.xml not found`);
  return new TextDecoder().decode(part.data);
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
});
