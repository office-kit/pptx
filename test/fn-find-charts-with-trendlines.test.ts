// `findChartsWithTrendlines(slide)` — slide-scoped trendline auditor.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideChart,
  findChartsWithTrendlines,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findChartsWithTrendlines', () => {
  it('returns charts with a trendline on any series', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      spec: {
        kind: 'line',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [1, 2], trendline: { type: 'linear' } }],
      },
    });
    addSlideChart(slide, {
      x: inches(3),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      spec: {
        kind: 'line',
        categories: ['A', 'B'],
        series: [{ name: 'Y', values: [3, 4] }],
      },
    });
    expect(findChartsWithTrendlines(slide).length).toBe(1);
  });

  it('returns an empty array when no chart has a trendline', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      spec: {
        kind: 'column',
        categories: ['A'],
        series: [{ name: 'X', values: [1] }],
      },
    });
    expect(findChartsWithTrendlines(slide)).toEqual([]);
  });
});
