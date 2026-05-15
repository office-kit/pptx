// getShapeChartSpec — fetch a single shape's ChartSpec without
// iterating getSlideCharts.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideChart,
  addSlideShape,
  getShapeChartSpec,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapeChartSpec', () => {
  it('returns the spec for a chart shape', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const chart = addSlideChart(slide, {
      x: inches(0), y: inches(0), w: inches(4), h: inches(3),
      spec: {
        kind: 'bar',
        categories: ['Q1', 'Q2', 'Q3'],
        series: [{ name: 'Revenue', values: [10, 20, 30] }],
      },
    });
    const spec = getShapeChartSpec(chart);
    expect(spec).not.toBeNull();
    expect(spec!.kind).toBe('bar');
    expect(spec!.categories).toEqual(['Q1', 'Q2', 'Q3']);
    expect(spec!.series).toHaveLength(1);
    expect(spec!.series[0]!.name).toBe('Revenue');
  });

  it('returns null for a non-chart shape', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const rect = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    expect(getShapeChartSpec(rect)).toBeNull();
  });
});
