// ChartTextStyle.font — author a font face on chart labels and round-trip
// it back. The builder writes both the latin and east-asian typeface
// slots so CJK families (e.g. "Yu Gothic") aren't dropped to a latin-only
// fallback; the reader recovers the face from the latin slot.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  _internalPackageOf,
  addSlideChart,
  getSlideCharts,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setChartSpec,
} from '../src/api/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const decode = (b: Uint8Array): string => new TextDecoder().decode(b);
const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

const YU_GOTHIC = 'Yu Gothic';

describe('fn API: ChartTextStyle.font', () => {
  it('round-trips font on title / axis / legend / data-label styles', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['一月', '二月'],
        series: [{ name: '売上', values: [10, 20] }],
        title: '四半期',
        titleStyle: { font: YU_GOTHIC, sizePt: 18 },
        categoryAxisLabelStyle: { font: YU_GOTHIC },
        valueAxisLabelStyle: { font: YU_GOTHIC, color: '#333333' },
        dataLabels: {
          showValue: true,
          showCategory: false,
          showSeriesName: false,
          showPercent: false,
          textStyle: { font: YU_GOTHIC },
        },
        legend: { position: 'r', textStyle: { font: YU_GOTHIC } },
      },
    });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.titleStyle?.font).toBe(YU_GOTHIC);
    expect(spec.titleStyle?.sizePt).toBe(18);
    expect(spec.categoryAxisLabelStyle?.font).toBe(YU_GOTHIC);
    expect(spec.valueAxisLabelStyle?.font).toBe(YU_GOTHIC);
    expect(spec.valueAxisLabelStyle?.color).toBe('#333333');
    expect(spec.dataLabels?.textStyle?.font).toBe(YU_GOTHIC);
    expect(spec.legend?.textStyle?.font).toBe(YU_GOTHIC);
  });

  it('emits both <a:latin> and <a:ea> typeface for the same face', async () => {
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
        series: [{ name: 'X', values: [1, 2] }],
        title: 'T',
        titleStyle: { font: YU_GOTHIC },
      },
    });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const pkg = _internalPackageOf(reloaded);
    const chartPart = pkg.parts.find((p) => p.name === '/ppt/charts/chart1.xml');
    const xml = decode(chartPart!.data);
    expect(xml).toContain(`<a:latin typeface="${YU_GOTHIC}"`);
    expect(xml).toContain(`<a:ea typeface="${YU_GOTHIC}"`);
  });

  it('applies font on setChartSpec updates too', async () => {
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
        series: [{ name: 'X', values: [1, 2] }],
      },
    });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const chart = getSlideCharts(getSlides(reloaded)[0]!)[0]!;
    setChartSpec(chart, {
      kind: 'column',
      categories: ['A', 'B'],
      series: [{ name: 'X', values: [1, 2] }],
      title: 'Styled',
      titleStyle: { font: YU_GOTHIC },
    });
    const reread = await loadPresentation(await savePresentation(reloaded));
    const spec = getSlideCharts(getSlides(reread)[0]!)[0]!.spec!;
    expect(spec.titleStyle?.font).toBe(YU_GOTHIC);
  });

  it('omits the typeface children entirely when no font is authored', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(3),
      spec: {
        kind: 'bar',
        categories: ['A'],
        series: [{ name: 'X', values: [1] }],
        title: 'NoFont',
        titleStyle: { sizePt: 12 },
      },
    });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.titleStyle?.font).toBeUndefined();
    const pkg = _internalPackageOf(reloaded);
    const xml = decode(pkg.parts.find((p) => p.name === '/ppt/charts/chart1.xml')!.data);
    expect(xml).not.toContain('<a:latin');
  });

  skipIfNoXmllint('a chart with authored fonts stays schema-valid', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['一月', '二月'],
        series: [{ name: '売上', values: [10, 20] }],
        title: '四半期',
        titleStyle: { font: YU_GOTHIC, sizePt: 18, color: '#FF0000' },
        categoryAxisLabelStyle: { font: YU_GOTHIC },
        valueAxisLabelStyle: { font: YU_GOTHIC },
        legend: { position: 'b', textStyle: { font: YU_GOTHIC } },
      },
    });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const pkg = _internalPackageOf(reloaded);
    const chartPart = pkg.parts.find((p) => p.name === '/ppt/charts/chart1.xml');
    expectSchemaValid(decode(chartPart!.data), 'chart');
  });
});
