// `getPresentationChartKindCounts(pres)` — deck-wide chart kind
// histogram. Companion to `findSlidesWithChartKind`.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideChart,
  getPresentationChartKindCounts,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getPresentationChartKindCounts', () => {
  it('returns all-zeros for a deck with no charts', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getPresentationChartKindCounts(pres)).toEqual({
      bar: 0,
      column: 0,
      line: 0,
      pie: 0,
      doughnut: 0,
      area: 0,
      scatter: 0,
      radar: 0,
      bubble: 0,
    });
  });

  it('counts every chart across every slide', async () => {
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
      x: inches(3),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      spec: { kind: 'pie', categories: ['A'], series: [{ name: 'X', values: [1] }] },
    });
    addSlideChart(slideB!, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      spec: { kind: 'column', categories: ['A'], series: [{ name: 'X', values: [1] }] },
    });
    expect(getPresentationChartKindCounts(pres)).toEqual({
      bar: 0,
      column: 2,
      line: 0,
      pie: 1,
      doughnut: 0,
      area: 0,
      scatter: 0,
      radar: 0,
      bubble: 0,
    });
  });

  it('always includes every kind in the result (even zeros)', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const slide = addBlankSlide(pres);
    addSlideChart(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      spec: { kind: 'line', categories: ['A'], series: [{ name: 'X', values: [1] }] },
    });
    const counts = getPresentationChartKindCounts(pres);
    expect(Object.keys(counts).sort()).toEqual(
      ['area', 'bar', 'bubble', 'column', 'doughnut', 'line', 'pie', 'radar', 'scatter'].sort(),
    );
    expect(counts.line).toBe(1);
  });
});
