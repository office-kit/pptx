// `ChartSeries.pointDataLabels` — per-point `<c:dLbl>` overrides, the
// mechanism pie / doughnut exporters use to give each slice its own label
// content and font.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  type ChartDataLabels,
  type ChartSpec,
  addSlideChart,
  getSlideCharts,
  getSlides,
  inches,
  loadPresentation,
  readPackagePart,
  savePresentation,
} from '../src/api/index.ts';
import { readChartSpec } from '../src/internal/chartml/index.ts';
import { parseXml } from '../src/internal/xml/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const decoder = new TextDecoder();

const roundTrip = async (
  spec: ChartSpec,
): Promise<{ readonly spec: ChartSpec; readonly xml: string }> => {
  const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
  addSlideChart(getSlides(pres)[0]!, {
    x: inches(0.5),
    y: inches(0.5),
    w: inches(6),
    h: inches(4),
    spec,
  });
  const reloaded = await loadPresentation(await savePresentation(pres));
  const xml = decoder.decode(readPackagePart(reloaded, '/ppt/charts/chart1.xml')!);
  return { spec: getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!, xml };
};

const valueAndPercent: ChartDataLabels = {
  showValue: true,
  showCategory: false,
  showSeriesName: false,
  showPercent: true,
  numberFormat: '0%',
  textStyle: { font: 'Yu Gothic', sizePt: 12 },
};

describe('fn API: pointDataLabels', () => {
  it('rejects per-point overrides without the series-level labels they fall back to', async () => {
    await expect(
      roundTrip({
        kind: 'doughnut',
        categories: ['A', 'B', 'C'],
        series: [
          { name: 'Share', values: [50, 30, 20], pointDataLabels: [valueAndPercent, null, null] },
        ],
      }),
    ).rejects.toThrow(/pointDataLabels without dataLabels/);
  });

  skipIfNoXmllint('round-trips sparse per-point overrides on a doughnut', async () => {
    const seriesLabels: ChartDataLabels = {
      showValue: false,
      showCategory: true,
      showSeriesName: false,
      showPercent: false,
      textStyle: { font: 'Arial', sizePt: 18 },
    };
    const { spec, xml } = await roundTrip({
      kind: 'doughnut',
      categories: ['A', 'B', 'C'],
      series: [
        {
          name: 'Share',
          values: [50, 30, 20],
          dataLabels: seriesLabels,
          pointDataLabels: [valueAndPercent, null, valueAndPercent],
        },
      ],
    });
    expect(spec.series[0]!.pointDataLabels).toEqual([valueAndPercent, null, valueAndPercent]);
    expect(spec.series[0]!.dataLabels).toEqual(seriesLabels);
    expect(xml).toContain('<c:dLbls><c:dLbl><c:idx val="0"/>');
    expect(xml).toContain('<c:dLbl><c:idx val="2"/>');
    expect(xml).not.toContain('<c:dLbl><c:idx val="1"/>');
    expectSchemaValid(xml, 'chart');
  });

  skipIfNoXmllint('round-trips per-point overrides with a position on a column chart', async () => {
    const outEnd: ChartDataLabels = { ...valueAndPercent, showPercent: false, position: 'outEnd' };
    const seriesLabels: ChartDataLabels = {
      showValue: true,
      showCategory: false,
      showSeriesName: false,
      showPercent: false,
    };
    const { spec, xml } = await roundTrip({
      kind: 'column',
      categories: ['A', 'B'],
      series: [
        { name: 'Count', values: [1, 2], dataLabels: seriesLabels, pointDataLabels: [outEnd] },
      ],
    });
    expect(spec.series[0]!.pointDataLabels).toEqual([outEnd]);
    expect(xml).toContain('<c:dLbl><c:idx val="0"/><c:numFmt formatCode="0%" sourceLinked="0"/>');
    expectSchemaValid(xml, 'chart');
  });

  it('round-trips per-point overrides alongside series-level labels', async () => {
    const seriesLabels: ChartDataLabels = {
      showValue: false,
      showCategory: true,
      showSeriesName: false,
      showPercent: false,
      textStyle: { font: 'Arial', sizePt: 18 },
    };
    const { spec } = await roundTrip({
      kind: 'pie',
      categories: ['A', 'B'],
      series: [
        {
          name: 'Share',
          values: [60, 40],
          dataLabels: seriesLabels,
          pointDataLabels: [valueAndPercent, null],
        },
      ],
    });
    expect(spec.series[0]!.dataLabels).toEqual(seriesLabels);
    // Sparse like `pointColors`: trailing `null` slots are not read back.
    expect(spec.series[0]!.pointDataLabels).toEqual([valueAndPercent]);
  });
});

describe('chart reader: per-point data labels', () => {
  const wrap = (ser: string): string => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart><c:plotArea><c:layout/><c:doughnutChart><c:varyColors val="1"/>${ser}</c:doughnutChart></c:plotArea></c:chart>
</c:chartSpace>`;

  it('reads a deleted point label as an override with every toggle off', () => {
    const spec = readChartSpec(
      parseXml(
        wrap(`
      <c:ser>
        <c:idx val="0"/><c:order val="0"/>
        <c:tx><c:strLit><c:pt idx="0"><c:v>Share</c:v></c:pt></c:strLit></c:tx>
        <c:dLbls>
          <c:dLbl><c:idx val="0"/><c:delete val="1"/></c:dLbl>
          <c:dLbl><c:idx val="1"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbl>
          <c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="1"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>
        </c:dLbls>
        <c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val>
      </c:ser>`),
      ).root,
    )!;
    expect(spec.series[0]!.pointDataLabels).toEqual([
      { showValue: false, showCategory: false, showSeriesName: false, showPercent: false },
      { showValue: true, showCategory: false, showSeriesName: false, showPercent: false },
    ]);
    expect(spec.series[0]!.dataLabels).toEqual({
      showValue: false,
      showCategory: true,
      showSeriesName: false,
      showPercent: false,
    });
  });

  const readDoughnut = (ser: string): ChartSpec => readChartSpec(parseXml(wrap(ser)).root)!;
  const twoPoints = `<c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val>`;
  const seriesShowVal = `<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>`;

  it('drops a point label whose idx is past the series point count', () => {
    const spec = readDoughnut(`
      <c:ser>
        <c:idx val="0"/><c:order val="0"/>
        <c:dPt><c:idx val="100000000"/><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr></c:dPt>
        <c:dLbls>
          <c:dLbl><c:idx val="100000000"/>${seriesShowVal}</c:dLbl>
          <c:dLbl><c:idx val="99999999999999999999"/>${seriesShowVal}</c:dLbl>
          <c:dLbl><c:idx val="1e1"/>${seriesShowVal}</c:dLbl>
          <c:dLbl><c:idx val="2"/>${seriesShowVal}</c:dLbl>
          ${seriesShowVal}
        </c:dLbls>
        ${twoPoints}
      </c:ser>`);
    expect(spec.series[0]!.pointDataLabels).toBeUndefined();
    expect(spec.series[0]!.pointColors).toBeUndefined();
  });

  it('inherits the toggles a point label leaves out from the series group', () => {
    const spec = readDoughnut(`
      <c:ser>
        <c:idx val="0"/><c:order val="0"/>
        <c:dLbls>
          <c:dLbl><c:idx val="0"/><c:numFmt formatCode="0.0" sourceLinked="0"/></c:dLbl>
          ${seriesShowVal}
        </c:dLbls>
        ${twoPoints}
      </c:ser>`);
    expect(spec.series[0]!.pointDataLabels).toEqual([
      {
        showValue: true,
        showCategory: false,
        showSeriesName: false,
        showPercent: false,
        numberFormat: '0.0',
      },
    ]);
  });

  it('reads the toggles as false when the series group is missing', () => {
    const spec = readDoughnut(`
      <c:ser>
        <c:idx val="0"/><c:order val="0"/>
        <c:dLbls><c:dLbl><c:idx val="1"/><c:showPercent val="true"/></c:dLbl></c:dLbls>
        ${twoPoints}
      </c:ser>`);
    expect(spec.series[0]!.dataLabels).toBeUndefined();
    expect(spec.series[0]!.pointDataLabels).toEqual([
      null,
      { showValue: false, showCategory: false, showSeriesName: false, showPercent: true },
    ]);
  });

  it('turns a deleted point label off even when the series shows values', () => {
    const spec = readDoughnut(`
      <c:ser>
        <c:idx val="0"/><c:order val="0"/>
        <c:dLbls>
          <c:dLbl><c:idx val="0"/><c:delete val="1"/></c:dLbl>
          <c:dLbl><c:idx val="1"/><c:delete/></c:dLbl>
          ${seriesShowVal}
        </c:dLbls>
        ${twoPoints}
      </c:ser>`);
    expect(spec.series[0]!.pointDataLabels).toEqual([
      { showValue: false, showCategory: false, showSeriesName: false, showPercent: false },
      { showValue: false, showCategory: false, showSeriesName: false, showPercent: false },
    ]);
  });

  skipIfNoXmllint('round-trips partial and deleted point labels through the builder', async () => {
    const read = readDoughnut(`
      <c:ser>
        <c:idx val="0"/><c:order val="0"/>
        <c:tx><c:strLit><c:pt idx="0"><c:v>Share</c:v></c:pt></c:strLit></c:tx>
        <c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
        <c:dLbls>
          <c:dLbl><c:idx val="0"/><c:numFmt formatCode="0.0" sourceLinked="0"/></c:dLbl>
          <c:dLbl><c:idx val="1"/><c:delete val="1"/></c:dLbl>
          ${seriesShowVal}
        </c:dLbls>
        ${twoPoints}
      </c:ser>`);
    const { spec, xml } = await roundTrip(read);
    expect(spec.series[0]!.pointDataLabels).toEqual(read.series[0]!.pointDataLabels);
    expect(spec.series[0]!.dataLabels).toEqual(read.series[0]!.dataLabels);
    expectSchemaValid(xml, 'chart');
  });

  skipIfNoXmllint(
    'round-trips showLeaderLines on series and chart labels, never per point',
    async () => {
      const labels = { ...valueAndPercent, separator: '; ', showLeaderLines: false };
      const { spec, xml } = await roundTrip({
        kind: 'doughnut',
        categories: ['A', 'B'],
        series: [
          {
            name: 'Share',
            values: [3, 1],
            dataLabels: labels,
            pointDataLabels: [{ ...valueAndPercent, showLeaderLines: true }, null],
          },
        ],
        dataLabels: { ...valueAndPercent, showLeaderLines: true },
      });
      const dLbls = xml.slice(xml.indexOf('<c:dLbls>'), xml.indexOf('</c:dLbls>'));
      const point = dLbls.slice(0, dLbls.indexOf('</c:dLbl>'));
      expect(point).not.toContain('showLeaderLines');
      expect(dLbls).toContain('<c:separator>; </c:separator><c:showLeaderLines val="0"/>');
      expect(xml).toContain('<c:showBubbleSize val="0"/><c:showLeaderLines val="1"/>');
      expect(spec.series[0]!.dataLabels).toEqual(labels);
      expect(spec.series[0]!.pointDataLabels).toEqual([valueAndPercent]);
      expect(spec.dataLabels?.showLeaderLines).toBe(true);
      expectSchemaValid(xml, 'chart');
    },
  );

  it('reads an absent showLeaderLines as unset', () => {
    const spec = readDoughnut(`
      <c:ser>
        <c:idx val="0"/><c:order val="0"/>
        <c:dLbls>${seriesShowVal}</c:dLbls>
        ${twoPoints}
      </c:ser>`);
    expect(spec.series[0]!.dataLabels?.showLeaderLines).toBeUndefined();
  });

  it('reads PptxGenJS-shaped point labels (spPr inside <c:dLbl>, group on the parent)', () => {
    const spec = readChartSpec(
      parseXml(
        wrap(`
      <c:ser>
        <c:idx val="0"/><c:order val="0"/>
        <c:tx><c:strLit><c:pt idx="0"><c:v>Share</c:v></c:pt></c:strLit></c:tx>
        <c:dLbls>
          <c:dLbl><c:idx val="0"/><c:spPr/><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200"><a:latin typeface="Yu Gothic"/></a:defRPr></a:pPr></a:p></c:txPr><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/></c:dLbl>
          <c:numFmt formatCode="0%" sourceLinked="0"/>
          <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1800"><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr>
          <c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="1"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>
        </c:dLbls>
        <c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
      </c:ser>`),
      ).root,
    )!;
    expect(spec.series[0]!.pointDataLabels).toEqual([
      {
        showValue: true,
        showCategory: false,
        showSeriesName: false,
        showPercent: true,
        textStyle: { font: 'Yu Gothic', sizePt: 12 },
      },
    ]);
    expect(spec.series[0]!.dataLabels).toEqual({
      showValue: false,
      showCategory: true,
      showSeriesName: false,
      showPercent: false,
      numberFormat: '0%',
      textStyle: { font: 'Arial', sizePt: 18 },
    });
  });

  it('pads values to <c:ptCount> when empty points are omitted', () => {
    const spec = readChartSpec(
      parseXml(
        wrap(`
      <c:ser>
        <c:idx val="0"/><c:order val="0"/>
        <c:tx><c:strLit><c:pt idx="0"><c:v>Share</c:v></c:pt></c:strLit></c:tx>
        <c:val><c:numLit><c:ptCount val="3"/><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val>
      </c:ser>`),
      ).root,
    )!;
    expect(spec.series[0]!.values).toEqual([null, 2, null]);
  });
});
