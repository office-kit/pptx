// getShapeEffect — read-back parity for setShapeShadow / setShapeGlow.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  clearShapeEffects,
  getShapeEffect,
  getSlides,
  inches,
  loadPresentation,
  setShapeGlow,
  setShapeShadow,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapeEffect', () => {
  it('returns null when no effect is set', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(2), h: inches(2),
    });
    expect(getShapeEffect(shape)).toBeNull();
  });

  it('round-trips a shadow', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(2), h: inches(2),
    });
    setShapeShadow(shape, { color: '#112233', angleDeg: 45, opacity: 0.5 });
    const got = getShapeEffect(shape);
    expect(got?.kind).toBe('shadow');
    if (got?.kind === 'shadow') {
      expect(got.color).toBe('#112233');
      expect(got.angleDeg).toBeCloseTo(45);
      expect(got.opacity).toBeCloseTo(0.5);
    }
  });

  it('round-trips a glow', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'ellipse', x: inches(0), y: inches(0), w: inches(2), h: inches(2),
    });
    setShapeGlow(shape, { color: '#FF0000', radiusEmu: 80000 });
    const got = getShapeEffect(shape);
    expect(got?.kind).toBe('glow');
    if (got?.kind === 'glow') {
      expect(got.color).toBe('#FF0000');
      expect(got.radiusEmu).toBe(80000);
    }
  });

  it('returns null after clearShapeEffects', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(2), h: inches(2),
    });
    setShapeShadow(shape);
    clearShapeEffects(shape);
    expect(getShapeEffect(shape)).toBeNull();
  });
});
