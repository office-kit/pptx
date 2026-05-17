// getShapeChartCategories / getShapeChartSeriesNames — quick
// chart-shape introspection.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideChart,
  addSlideShape,
  getShapeChartCategories,
  getShapeChartSeriesNames,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: chart axis readers', () => {
  it('returns categories + series names', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const chart = addSlideChart(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(3),
      spec: {
        kind: 'bar',
        categories: ['Q1', 'Q2', 'Q3'],
        series: [
          { name: 'Revenue', values: [10, 20, 30] },
          { name: 'Costs', values: [5, 7, 9] },
        ],
      },
    });
    expect(getShapeChartCategories(chart)).toEqual(['Q1', 'Q2', 'Q3']);
    expect(getShapeChartSeriesNames(chart)).toEqual(['Revenue', 'Costs']);
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
    expect(getShapeChartCategories(rect)).toBeNull();
    expect(getShapeChartSeriesNames(rect)).toBeNull();
  });
});
