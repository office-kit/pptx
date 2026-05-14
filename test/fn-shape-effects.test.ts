// Shape effects: shadow + glow.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  Presentation,
  _internalPackageOf,
  addSlideShape,
  clearShapeEffects,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeGlow,
  setShapeShadow,
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

describe('fn API: shape effects', () => {
  it('setShapeShadow writes an outerShdw with computed direction', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(2), h: inches(2),
    });
    setShapeShadow(shape, { color: '#000000', angleDeg: 45, opacity: 0.5 });
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('<a:effectLst>');
    expect(xml).toContain('<a:outerShdw');
    // 45° × 60000 = 2700000
    expect(xml).toContain('dir="2700000"');
    // 0.5 opacity → alpha 50000
    expect(xml).toContain('val="50000"');
  });

  it('setShapeGlow writes a glow with the configured radius', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'ellipse', x: inches(0), y: inches(0), w: inches(2), h: inches(2),
    });
    setShapeGlow(shape, { color: '#FF0000', radiusEmu: 90000 });
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('<a:glow');
    expect(xml).toContain('rad="90000"');
    expect(xml).toContain('FF0000');
  });

  it('shadow then glow replaces the prior effect list', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(2), h: inches(2),
    });
    setShapeShadow(shape);
    setShapeGlow(shape, { color: '#00FF00' });
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('<a:glow');
    expect(xml).not.toContain('<a:outerShdw');
  });

  it('clearShapeEffects removes any effectLst', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(2), h: inches(2),
    });
    setShapeShadow(shape);
    expect(await slideXml(await savePresentation(pres), 0)).toContain('<a:effectLst>');
    clearShapeEffects(shape);
    expect(await slideXml(await savePresentation(pres), 0)).not.toContain('<a:effectLst>');
  });
});
