// `getSlideCharts` — parse charts back from a deck.
//
// Round-trip the same `ChartSpec` through `addSlideChart` →
// `savePresentation` → `loadPresentation` → `getSlideCharts` and
// verify the spec survives.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideChart,
  getSlideCharts,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideCharts', () => {
  it('round-trips a column chart spec', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['Q1', 'Q2', 'Q3', 'Q4'],
        series: [
          { name: 'Revenue', values: [10, 20, 15, 30] },
          { name: 'Cost', values: [5, 7, 9, 11] },
        ],
        title: 'Quarterly',
      },
    });

    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const charts = getSlideCharts(getSlides(reloaded)[0]!);
    expect(charts).toHaveLength(1);
    const spec = charts[0]!.spec;
    expect(spec).not.toBeNull();
    expect(spec!.kind).toBe('column');
    expect(spec!.categories).toEqual(['Q1', 'Q2', 'Q3', 'Q4']);
    expect(spec!.series.map((s) => s.name)).toEqual(['Revenue', 'Cost']);
    expect(spec!.series[0]!.values).toEqual([10, 20, 15, 30]);
    expect(spec!.series[1]!.values).toEqual([5, 7, 9, 11]);
    expect(spec!.title).toBe('Quarterly');
  });

  it('round-trips titleStyle through builder', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [1, 2] }],
        title: 'Styled',
        titleStyle: { sizePt: 18, bold: true, color: '#FF0000' },
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const charts = getSlideCharts(getSlides(reloaded)[0]!);
    const spec = charts[0]!.spec!;
    expect(spec.title).toBe('Styled');
    expect(spec.titleStyle?.sizePt).toBe(18);
    expect(spec.titleStyle?.bold).toBe(true);
    expect(spec.titleStyle?.color).toBe('#FF0000');
  });

  it('distinguishes bar from column on the same barChart wire format', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(3),
      spec: {
        kind: 'bar',
        categories: ['A', 'B'],
        series: [{ name: 'S', values: [1, 2] }],
      },
    });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.kind).toBe('bar');
  });

  it('reads pie and doughnut kinds', async () => {
    for (const kind of ['pie', 'doughnut'] as const) {
      const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
      const slide = getSlides(pres)[0]!;
      addSlideChart(slide, {
        x: inches(0),
        y: inches(0),
        w: inches(4),
        h: inches(3),
        spec: {
          kind,
          categories: ['X', 'Y', 'Z'],
          series: [{ name: 'S', values: [1, 2, 3] }],
        },
      });
      const reloaded = await loadPresentation(await savePresentation(pres));
      expect(getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!.kind).toBe(kind);
    }
  });

  it('returns custom series colors verbatim', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(3),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [1, 2], color: '#112233' }],
      },
    });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.series[0]!.color).toBe('#112233');
  });

  it('returns empty array for slides without charts', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    expect(getSlideCharts(slide)).toEqual([]);
  });

  it('parses an unmodified template chart loaded fresh from disk', async () => {
    // Add a chart, save, load again — proves the reader doesn't depend
    // on in-memory state created by addSlideChart.
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    addSlideChart(getSlides(pres)[0]!, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(3),
      spec: {
        kind: 'line',
        categories: ['Jan', 'Feb', 'Mar'],
        series: [{ name: 'A', values: [3, 5, 7] }],
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const charts = getSlideCharts(getSlides(reloaded)[0]!);
    expect(charts[0]!.spec!.kind).toBe('line');
  });
});
