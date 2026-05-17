// `getShapeEffects` — returns every effect on the shape's `<a:effectLst>`
// in document order, not just the first one. Renderers need the full
// list because PowerPoint composes shadow + glow + softEdge into a
// single filter stack.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  getShapeEffects,
  getSlides,
  inches,
  loadPresentation,
  setShapeGlow,
  setShapeShadow,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapeEffects', () => {
  it('returns an empty array when no effects are set', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
    });
    expect(getShapeEffects(pres, shape)).toEqual([]);
  });

  it('reads the outer shadow set by setShapeShadow with all numeric fields', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
    });
    setShapeShadow(shape, { color: '#000000', angleDeg: 45, opacity: 0.5 });
    const effects = getShapeEffects(pres, shape);
    expect(effects).toHaveLength(1);
    expect(effects[0]!.kind).toBe('outerShdw');
    if (effects[0]!.kind === 'outerShdw') {
      expect(effects[0]!.color).toBe('#000000');
      expect(effects[0]!.angleDeg).toBeCloseTo(45);
      expect(effects[0]!.opacity).toBeCloseTo(0.5, 3);
      expect(effects[0]!.blurEmu).toBeGreaterThan(0);
    }
  });

  it('reads the glow set by setShapeGlow', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
    });
    setShapeGlow(shape, { color: '#FF0000', radiusEmu: 63500 });
    const effects = getShapeEffects(pres, shape);
    expect(effects).toHaveLength(1);
    expect(effects[0]!.kind).toBe('glow');
    if (effects[0]!.kind === 'glow') {
      expect(effects[0]!.color).toBe('#FF0000');
      expect(effects[0]!.radiusEmu).toBe(63500);
    }
  });
});
