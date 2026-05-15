// getSlidesWithCharts / getSlidesWithTables — sibling filters.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideChart,
  addSlideTable,
  getSlideIndex,
  getSlides,
  getSlidesWithCharts,
  getSlidesWithTables,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlidesWithCharts / getSlidesWithTables', () => {
  it('lists each slide that has a chart', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    addSlideChart(getSlides(pres)[0]!, {
      x: inches(0), y: inches(0), w: inches(3), h: inches(2),
      spec: { kind: 'bar', categories: ['Q1'], series: [{ name: 'r', values: [1] }] },
    });
    const hits = getSlidesWithCharts(pres);
    expect(hits.length).toBe(1);
    expect(getSlideIndex(pres, hits[0]!)).toBe(0);
  });

  it('lists each slide that has a table', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    addSlideTable(getSlides(pres)[1]!, {
      x: inches(0), y: inches(0), w: inches(3), h: inches(2),
      rows: [['a', 'b']],
    });
    const hits = getSlidesWithTables(pres);
    expect(hits.length).toBe(1);
    expect(getSlideIndex(pres, hits[0]!)).toBe(1);
  });

  it('both return empty on a fresh deck', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getSlidesWithCharts(pres)).toEqual([]);
    expect(getSlidesWithTables(pres)).toEqual([]);
  });
});
