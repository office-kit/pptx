// `findSlidesWithChartTrendlines(pres)` — deck-level trendline auditor.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideChart,
  findSlidesWithChartTrendlines,
  getSlideIndex,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findSlidesWithChartTrendlines', () => {
  it('returns every slide with at least one trendline-carrying chart', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [slideA, slideB] = getSlides(pres);
    addSlideChart(slideA!, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      spec: {
        kind: 'line',
        categories: ['A'],
        series: [{ name: 'X', values: [1], trendline: { type: 'linear' } }],
      },
    });
    addSlideChart(slideB!, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      spec: {
        kind: 'line',
        categories: ['A'],
        series: [{ name: 'Y', values: [2] }],
      },
    });
    const slides = findSlidesWithChartTrendlines(pres);
    expect(slides.length).toBe(1);
    expect(getSlideIndex(pres, slides[0]!)).toBe(getSlideIndex(pres, slideA!));
  });

  it('returns an empty array when no chart has a trendline', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(findSlidesWithChartTrendlines(pres)).toEqual([]);
  });
});
