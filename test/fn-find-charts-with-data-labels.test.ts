// `findChartsWithDataLabels(slide)` — slide-scoped audit for charts
// that show values / categories / percentages on data points.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideChart,
  findChartsWithDataLabels,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findChartsWithDataLabels', () => {
  it('matches charts with chart-level dataLabels', async () => {
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
        dataLabels: {
          showValue: true,
          showCategory: false,
          showSeriesName: false,
          showPercent: false,
        },
      },
    });
    expect(findChartsWithDataLabels(slide).length).toBe(1);
  });

  it('matches charts with only per-series dataLabels', async () => {
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
        series: [
          {
            name: 'X',
            values: [1],
            dataLabels: {
              showValue: true,
              showCategory: false,
              showSeriesName: false,
              showPercent: false,
            },
          },
        ],
      },
    });
    expect(findChartsWithDataLabels(slide).length).toBe(1);
  });

  it('returns an empty array when no chart enables a show* flag', async () => {
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
    expect(findChartsWithDataLabels(slide)).toEqual([]);
  });

  it('ignores a dataLabels object whose every show* flag is false', async () => {
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
        dataLabels: {
          showValue: false,
          showCategory: false,
          showSeriesName: false,
          showPercent: false,
        },
      },
    });
    expect(findChartsWithDataLabels(slide)).toEqual([]);
  });
});
