// getAllCharts — every chart paired with its slide index.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideChart,
  getAllCharts,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getAllCharts', () => {
  it('returns an empty list when no slides have charts', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getAllCharts(pres)).toEqual([]);
  });

  it('pairs each chart with the slide it lives on', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [first, second] = getSlides(pres);
    addSlideChart(first!, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      spec: { kind: 'bar', categories: ['Q1'], series: [{ name: 'r', values: [1] }] },
    });
    addSlideChart(second!, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      spec: { kind: 'line', categories: ['Q1'], series: [{ name: 'r', values: [2] }] },
    });
    addSlideChart(second!, {
      x: inches(4),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      spec: { kind: 'pie', categories: ['A'], series: [{ name: 'r', values: [3] }] },
    });

    const entries = getAllCharts(pres);
    expect(entries.length).toBe(3);
    expect(entries.map((e) => e.slideIndex)).toEqual([0, 1, 1]);
    expect(entries.map((e) => e.chart.spec!.kind)).toEqual(['bar', 'line', 'pie']);
  });
});
