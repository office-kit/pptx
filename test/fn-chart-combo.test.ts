// Combo charts — per-series `chartKind` overrides and the secondary
// value axis (`secondaryAxis: true`).
//
// Verifies end-to-end:
//   - The builder splits the series into plot groups (`<c:barChart>` +
//     `<c:lineChart>`) and emits the secondary axis pair on demand.
//   - `getShapeChartSpec` round-trips `chartKind` / `secondaryAxis`.
//   - Non-combo base kinds reject the per-series fields.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideChart,
  getShapeChartSpec,
  getSlides,
  inches,
  loadPresentation,
  readPackagePart,
  savePresentation,
} from '../src/api/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const decoder = new TextDecoder();

describe('fn API: combo charts', () => {
  it('column + line(secondaryAxis) emits both plot groups and the secondary axis pair', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;

    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(8),
      h: inches(4.5),
      spec: {
        kind: 'column',
        categories: ['速い', '普通', '遅い'],
        series: [
          { name: '件数', values: [92, 118, 54] },
          { name: '平均スコア', values: [4.5, 3.8, 2.9], chartKind: 'line', secondaryAxis: true },
        ],
      },
    });

    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const chartXmlBytes = readPackagePart(reloaded, '/ppt/charts/chart1.xml');
    expect(chartXmlBytes).not.toBeNull();
    const xml = decoder.decode(chartXmlBytes!);

    expect(xml).toContain('<c:barChart>');
    expect(xml).toContain('<c:lineChart>');
    // Secondary value axis on the right, crossing at max, with its
    // deleted companion category axis.
    expect(xml).toContain('<c:axId val="444444444"/>');
    expect(xml).toContain('<c:axPos val="r"/>');
    expect(xml).toContain('<c:crosses val="max"/>');
    expect(xml).toContain('<c:axId val="333333333"/>');
    // The line group must reference the secondary pair, the bar group
    // the primary pair.
    const lineChartXml = xml.slice(xml.indexOf('<c:lineChart>'), xml.indexOf('</c:lineChart>'));
    expect(lineChartXml).toContain('<c:axId val="333333333"/>');
    expect(lineChartXml).toContain('<c:axId val="444444444"/>');
    const barChartXml = xml.slice(xml.indexOf('<c:barChart>'), xml.indexOf('</c:barChart>'));
    expect(barChartXml).toContain('<c:axId val="111111111"/>');
    expect(barChartXml).toContain('<c:axId val="222222222"/>');
  });

  it('round-trips chartKind / secondaryAxis through getShapeChartSpec', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;

    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(8),
      h: inches(4.5),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [
          { name: 'count', values: [100, 200] },
          { name: 'rate', values: [0.5, 0.7], chartKind: 'line', secondaryAxis: true },
        ],
      },
    });

    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const shapes = getSlides(reloaded)[0]!;
    const chartShape = (await import('../src/api/index.ts')).getSlideShapes(shapes).at(-1)!;
    const spec = getShapeChartSpec(chartShape);
    expect(spec).not.toBeNull();
    expect(spec!.kind).toBe('column');
    expect(spec!.series).toHaveLength(2);
    const [count, rate] = spec!.series;
    expect(count!.chartKind).toBeUndefined();
    expect(count!.secondaryAxis).toBeUndefined();
    expect(rate!.chartKind).toBe('line');
    expect(rate!.secondaryAxis).toBe(true);
    expect(rate!.values).toEqual([0.5, 0.7]);
  });

  it('rejects per-series combo fields on pie charts', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;

    expect(() =>
      addSlideChart(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(4),
        spec: {
          kind: 'pie',
          categories: ['A', 'B'],
          series: [{ name: 's', values: [1, 2], chartKind: 'line' }],
        },
      }),
    ).toThrow(/combo/);
  });
});

describe('fn API: combo plot groups', () => {
  const comboRoundTrip = async (spec: Parameters<typeof addSlideChart>[1]['spec']) => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    addSlideChart(getSlides(pres)[0]!, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(8),
      h: inches(4.5),
      spec,
    });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const chartShape = (await import('../src/api/index.ts'))
      .getSlideShapes(getSlides(reloaded)[0]!)
      .at(-1)!;
    const xml = decoder.decode(readPackagePart(reloaded, '/ppt/charts/chart1.xml')!);
    return { spec: getShapeChartSpec(chartShape)!, xml };
  };
  const between = (xml: string, open: string, close: string): string =>
    xml.slice(xml.indexOf(open), xml.indexOf(close, xml.indexOf(open)));

  skipIfNoXmllint('writes a line group as standard when the chart is clustered', async () => {
    const { spec, xml } = await comboRoundTrip({
      kind: 'column',
      categories: ['A', 'B'],
      grouping: 'clustered',
      series: [
        { name: 'count', values: [1, 2] },
        { name: 'rate', values: [0.5, 0.7], chartKind: 'line', secondaryAxis: true },
      ],
    });
    expect(between(xml, '<c:barChart>', '</c:barChart>')).toContain(
      '<c:grouping val="clustered"/>',
    );
    expect(between(xml, '<c:lineChart>', '</c:lineChart>')).toContain(
      '<c:grouping val="standard"/>',
    );
    expect(spec.grouping).toBe('clustered');
    expectSchemaValid(xml, 'chart');
  });

  skipIfNoXmllint('keeps a stacked grouping on an area group', async () => {
    const { xml } = await comboRoundTrip({
      kind: 'column',
      categories: ['A', 'B'],
      grouping: 'stacked',
      series: [
        { name: 'count', values: [1, 2] },
        { name: 'base', values: [3, 4], chartKind: 'area' },
      ],
    });
    expect(between(xml, '<c:areaChart>', '</c:areaChart>')).toContain(
      '<c:grouping val="stacked"/>',
    );
    expectSchemaValid(xml, 'chart');
  });

  skipIfNoXmllint('shapes a combo line series by its own kind (marker / smooth)', async () => {
    const { spec, xml } = await comboRoundTrip({
      kind: 'column',
      categories: ['A', 'B'],
      series: [
        { name: 'count', values: [1, 2] },
        {
          name: 'rate',
          values: [0.5, 0.7],
          chartKind: 'line',
          secondaryAxis: true,
          markerSymbol: 'circle',
          markerSizePt: 6,
          smooth: false,
        },
      ],
    });
    const lineSer = between(xml, '<c:lineChart>', '</c:lineChart>');
    expect(lineSer).toContain('<c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr>');
    expect(lineSer).toContain('<c:smooth val="0"/>');
    expect(between(xml, '<c:barChart>', '</c:barChart>')).not.toContain('<c:marker>');
    expect(spec.series[1]).toMatchObject({
      chartKind: 'line',
      markerSymbol: 'circle',
      markerSizePt: 6,
      smooth: false,
    });
    expect(spec.series[0]!.smooth).toBeUndefined();
    expectSchemaValid(xml, 'chart');
  });
});
