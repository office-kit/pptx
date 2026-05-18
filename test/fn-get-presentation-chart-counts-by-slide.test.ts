// `getPresentationChartCountsBySlide(pres)` — dense per-slide chart
// count array.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideChart,
  getPresentationChartCountsBySlide,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getPresentationChartCountsBySlide', () => {
  it('returns 0 for every slide on a chart-less deck', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getPresentationChartCountsBySlide(pres)).toEqual([0, 0]);
  });

  it('counts charts per slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [slideA, slideB] = getSlides(pres);
    addSlideChart(slideA!, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      spec: { kind: 'column', categories: ['A'], series: [{ name: 'X', values: [1] }] },
    });
    addSlideChart(slideA!, {
      x: inches(0),
      y: inches(2),
      w: inches(3),
      h: inches(2),
      spec: { kind: 'pie', categories: ['A'], series: [{ name: 'X', values: [1] }] },
    });
    addSlideChart(slideB!, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      spec: { kind: 'line', categories: ['A'], series: [{ name: 'X', values: [1] }] },
    });
    expect(getPresentationChartCountsBySlide(pres)).toEqual([2, 1]);
  });

  it('array length matches the slide count', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getPresentationChartCountsBySlide(pres).length).toBe(getSlides(pres).length);
  });
});
