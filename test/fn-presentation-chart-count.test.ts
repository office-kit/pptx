// getPresentationChartCount — fast counter for charts across the deck.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideChart,
  getAllCharts,
  getPresentationChartCount,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getPresentationChartCount', () => {
  it('matches getAllCharts length', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [first, second] = getSlides(pres);
    addSlideChart(first!, {
      x: inches(0), y: inches(0), w: inches(4), h: inches(3),
      spec: {
        kind: 'bar',
        categories: ['A', 'B'],
        series: [{ name: 'S', values: [1, 2] }],
      },
    });
    addSlideChart(second!, {
      x: inches(0), y: inches(0), w: inches(4), h: inches(3),
      spec: {
        kind: 'line',
        categories: ['X'],
        series: [{ name: 'T', values: [3] }],
      },
    });
    expect(getPresentationChartCount(pres)).toBe(getAllCharts(pres).length);
    expect(getPresentationChartCount(pres)).toBeGreaterThanOrEqual(2);
  });

  it('returns 0 on a deck with no charts', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getPresentationChartCount(pres)).toBe(0);
  });
});
