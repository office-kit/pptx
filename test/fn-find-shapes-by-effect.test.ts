// `findShapesByEffect(pres, slide, kind)` — visual-effect audit.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  findShapesByEffect,
  getSlides,
  inches,
  loadPresentation,
  setShapeGlow,
  setShapeShadow,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findShapesByEffect', () => {
  it('matches shapes by effect kind', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const a = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'A',
    });
    const b = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      text: 'B',
    });
    addSlideTextBox(slide, {
      x: inches(0),
      y: inches(2),
      w: inches(2),
      h: inches(1),
      text: 'C',
    });
    setShapeShadow(a, { color: '#000000', blurEmu: 50000, offsetEmu: 30000, angleDeg: 45 });
    setShapeGlow(b, { color: '#00FF00', radiusEmu: 50000 });
    expect(findShapesByEffect(pres, slide, 'outerShdw').length).toBe(1);
    expect(findShapesByEffect(pres, slide, 'glow').length).toBe(1);
    expect(findShapesByEffect(pres, slide, 'softEdge')).toEqual([]);
  });

  it('returns an empty array on a slide with no effects', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    expect(findShapesByEffect(pres, slide, 'glow')).toEqual([]);
  });
});
