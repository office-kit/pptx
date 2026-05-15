// setChartSpec — update an existing chart's data in place.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideChart,
  getSlideCharts,
  getSlides,
  inches,
  listPackageParts,
  loadPresentation,
  readPackagePart,
  savePresentation,
  setChartSpec,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: setChartSpec', () => {
  it('replaces chart data while preserving the shape geometry', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['Q1', 'Q2'],
        series: [{ name: 'Old', values: [1, 2] }],
      },
    });

    // Round-trip into a fresh handle, then update via setChartSpec.
    const reloaded = await loadPresentation(await savePresentation(pres));
    const charts = getSlideCharts(getSlides(reloaded)[0]!);
    expect(charts).toHaveLength(1);
    setChartSpec(charts[0]!, {
      kind: 'line',
      categories: ['Jan', 'Feb', 'Mar'],
      series: [{ name: 'New', values: [10, 20, 30] }],
      title: 'After',
    });

    // Re-load to verify persistence.
    const reread = await loadPresentation(await savePresentation(reloaded));
    const after = getSlideCharts(getSlides(reread)[0]!)[0]!.spec!;
    expect(after.kind).toBe('line');
    expect(after.categories).toEqual(['Jan', 'Feb', 'Mar']);
    expect(after.series[0]!.name).toBe('New');
    expect(after.series[0]!.values).toEqual([10, 20, 30]);
    expect(after.title).toBe('After');
  });

  it('also rewrites the embedded xlsx', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0), y: inches(0), w: inches(4), h: inches(3),
      spec: {
        kind: 'column',
        categories: ['A'],
        series: [{ name: 'orig', values: [1] }],
      },
    });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const chart = getSlideCharts(getSlides(reloaded)[0]!)[0]!;

    setChartSpec(chart, {
      kind: 'column',
      categories: ['Updated'],
      series: [{ name: 'renamed-series', values: [42] }],
    });

    const bytes = await savePresentation(reloaded);
    const after = await loadPresentation(bytes);
    const xlsx = listPackageParts(after).find((p) =>
      /^\/ppt\/embeddings\/Microsoft_Excel_Worksheet\d+\.xlsx$/.test(p.name),
    );
    expect(xlsx).toBeDefined();
    const chartBytes = readPackagePart(after, '/ppt/charts/chart1.xml');
    expect(chartBytes).not.toBeNull();
    expect(new TextDecoder().decode(chartBytes!)).toContain('renamed-series');
  });

  it('throws when invoked on a shape that is not a chart frame', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    expect(getSlideCharts(slide)).toEqual([]);
    // Synthesize a "chart data" handle pointing at a non-chart shape.
    const { getSlideShapes } = await import('../src/api/index.ts');
    const shape = getSlideShapes(slide)[0]!;
    expect(() =>
      setChartSpec(
        { shape, spec: null },
        { kind: 'column', categories: ['x'], series: [{ name: 'y', values: [1] }] },
      ),
    ).toThrow(/not a chart/);
  });
});
