// getShapeChartKind — quick discriminator over getShapeChartSpec.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideChart,
  addSlideShape,
  getShapeChartKind,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapeChartKind', () => {
  it('reports each kind authored via addSlideChart', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const bar = addSlideChart(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      spec: { kind: 'bar', categories: ['Q1'], series: [{ name: 'r', values: [1] }] },
    });
    const pie = addSlideChart(slide, {
      x: inches(4),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      spec: { kind: 'pie', categories: ['A'], series: [{ name: 'r', values: [1] }] },
    });

    expect(getShapeChartKind(bar)).toBe('bar');
    expect(getShapeChartKind(pie)).toBe('pie');
  });

  it('returns null for non-chart shapes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const rect = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    expect(getShapeChartKind(rect)).toBeNull();
  });
});
