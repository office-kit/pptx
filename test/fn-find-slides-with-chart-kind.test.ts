// `findSlidesWithChartKind` — kind-filtered version of
// `getSlidesWithCharts`.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideChart,
  findSlidesWithChartKind,
  getSlideIndex,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findSlidesWithChartKind', () => {
  it('returns only slides whose chart matches the kind', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [slideA, slideB] = getSlides(pres);
    addSlideChart(slideA!, {
      x: inches(0),
      y: inches(0),
      w: inches(6),
      h: inches(4),
      spec: { kind: 'column', categories: ['A'], series: [{ name: 'X', values: [1] }] },
    });
    addSlideChart(slideB!, {
      x: inches(0),
      y: inches(0),
      w: inches(6),
      h: inches(4),
      spec: { kind: 'pie', categories: ['A'], series: [{ name: 'X', values: [1] }] },
    });

    const cols = findSlidesWithChartKind(pres, 'column');
    expect(cols.map((s) => getSlideIndex(pres, s))).toEqual([getSlideIndex(pres, slideA!)]);
    const pies = findSlidesWithChartKind(pres, 'pie');
    expect(pies.map((s) => getSlideIndex(pres, s))).toEqual([getSlideIndex(pres, slideB!)]);
    expect(findSlidesWithChartKind(pres, 'line')).toEqual([]);
  });

  it('returns an empty array on a clean deck', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(findSlidesWithChartKind(pres, 'column')).toEqual([]);
  });

  it('counts a slide once even when it carries multiple charts of the same kind', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const slide = addBlankSlide(pres);
    for (let i = 0; i < 2; i++) {
      addSlideChart(slide, {
        x: inches(0),
        y: inches(i),
        w: inches(3),
        h: inches(1),
        spec: { kind: 'bar', categories: ['A'], series: [{ name: 'X', values: [1] }] },
      });
    }
    const bars = findSlidesWithChartKind(pres, 'bar');
    expect(bars).toHaveLength(1);
    expect(getSlideIndex(pres, bars[0]!)).toBe(getSlideIndex(pres, slide));
  });
});
