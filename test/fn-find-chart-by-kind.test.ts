// findChartByKind — locate a chart on a slide by its parsed kind.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideChart,
  findChartByKind,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findChartByKind', () => {
  it('finds the chart matching the requested kind', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      spec: { kind: 'bar', categories: ['Q1'], series: [{ name: 'r', values: [1] }] },
    });
    addSlideChart(slide, {
      x: inches(4),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      spec: { kind: 'line', categories: ['Q1'], series: [{ name: 'r', values: [1] }] },
    });

    const bar = findChartByKind(slide, 'bar');
    const line = findChartByKind(slide, 'line');
    expect(bar).not.toBeNull();
    expect(line).not.toBeNull();
    expect(bar!.spec!.kind).toBe('bar');
    expect(line!.spec!.kind).toBe('line');
  });

  it('returns null when no chart of that kind exists', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      spec: { kind: 'bar', categories: ['Q1'], series: [{ name: 'r', values: [1] }] },
    });
    expect(findChartByKind(slide, 'pie')).toBeNull();
  });

  it('returns null on a chart-free slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[1]!;
    expect(findChartByKind(slide, 'bar')).toBeNull();
  });
});
