// getPresentationSummary — diagnostic snapshot.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideChart,
  addSlideComment,
  addSlideShape,
  getPresentationSummary,
  getSlides,
  inches,
  loadPresentation,
  setShapeAnimation,
  setSlideHidden,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getPresentationSummary', () => {
  it('reports baseline counts for a clean fixture', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const s = getPresentationSummary(pres);
    expect(s.slideCount).toBe(2);
    expect(s.hiddenSlideCount).toBe(0);
    expect(s.layoutCount).toBeGreaterThan(0);
    expect(s.sectionCount).toBe(0);
    expect(s.hasCharts).toBe(false);
    expect(s.hasComments).toBe(false);
    expect(s.hasAnimations).toBe(false);
    expect(s.partCount).toBeGreaterThan(0);
    expect(typeof s.themeName).toBe('string');
  });

  it('flips hasCharts / hasComments / hasAnimations after mutations', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slides = getSlides(pres);
    addSlideChart(slides[0]!, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      spec: { kind: 'column', categories: ['A'], series: [{ name: 'S', values: [1] }] },
    });
    addSlideComment(slides[0]!, { author: { name: 'X' }, text: 'hi' });
    const sp = addSlideShape(slides[1]!, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    setShapeAnimation(sp, { effect: 'fadeIn' });
    setSlideHidden(slides[0]!, true);

    const s = getPresentationSummary(pres);
    expect(s.hasCharts).toBe(true);
    expect(s.hasComments).toBe(true);
    expect(s.hasAnimations).toBe(true);
    expect(s.hiddenSlideCount).toBe(1);
    expect(s.shapesByKind.graphicFrame).toBeGreaterThanOrEqual(1);
    expect(s.shapesByKind.shape).toBeGreaterThanOrEqual(1);
  });
});
