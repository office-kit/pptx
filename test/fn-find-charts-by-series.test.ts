// findChartsBySeriesName — locate every chart carrying a series of
// the given name.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideChart,
  findChartsBySeriesName,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findChartsBySeriesName', () => {
  it('matches charts whose series carry the name', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0), y: inches(0), w: inches(3), h: inches(2),
      spec: { kind: 'bar', categories: ['Q1'], series: [{ name: 'Revenue', values: [1] }] },
    });
    addSlideChart(slide, {
      x: inches(4), y: inches(0), w: inches(3), h: inches(2),
      spec: { kind: 'line', categories: ['Q1'], series: [{ name: 'Costs', values: [1] }] },
    });
    addSlideChart(slide, {
      x: inches(0), y: inches(3), w: inches(3), h: inches(2),
      spec: { kind: 'pie', categories: ['A'], series: [{ name: 'Revenue', values: [1] }] },
    });

    const revenue = findChartsBySeriesName(slide, 'Revenue');
    expect(revenue.length).toBe(2);
    for (const r of revenue) {
      expect(r.spec!.series.some((s) => s.name === 'Revenue')).toBe(true);
    }

    expect(findChartsBySeriesName(slide, 'Margin')).toEqual([]);
  });
});
