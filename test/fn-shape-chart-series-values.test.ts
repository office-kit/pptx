// getShapeChartSeriesValues — look up a single series's values by name.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideChart,
  addSlideShape,
  getShapeChartSeriesValues,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapeChartSeriesValues', () => {
  it('returns the named series values', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const chart = addSlideChart(slide, {
      x: inches(0), y: inches(0), w: inches(4), h: inches(3),
      spec: {
        kind: 'bar',
        categories: ['Q1', 'Q2'],
        series: [
          { name: 'Revenue', values: [10, 20] },
          { name: 'Costs', values: [5, 7] },
        ],
      },
    });
    expect(getShapeChartSeriesValues(chart, 'Revenue')).toEqual([10, 20]);
    expect(getShapeChartSeriesValues(chart, 'Costs')).toEqual([5, 7]);
  });

  it('returns null for an unknown series name', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const chart = addSlideChart(slide, {
      x: inches(0), y: inches(0), w: inches(4), h: inches(3),
      spec: { kind: 'bar', categories: ['Q1'], series: [{ name: 'r', values: [1] }] },
    });
    expect(getShapeChartSeriesValues(chart, 'no-such-series')).toBeNull();
  });

  it('returns null for non-chart shapes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const rect = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    expect(getShapeChartSeriesValues(rect, 'Revenue')).toBeNull();
  });
});
