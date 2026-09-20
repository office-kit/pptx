import { expect, it } from 'vitest';
import { unzipSync } from 'fflate';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';
import * as api from '../src/api/index.ts';
const box = { x: api.inches(1), y: api.inches(1), w: api.inches(5), h: api.inches(3) };
const series = { name: 'A', values: [1, 2] };
const labels = { showValue: true, showCategory: false, showSeriesName: false, showPercent: false };
const baseChart: api.ChartSpec = { kind: 'column', categories: ['a', 'b'], series: [series] };
const invalidCharts: ReadonlyArray<readonly [string, api.ChartSpec]> = [
  // @ts-expect-error Exercise the JavaScript boundary.
  ['kind', { ...baseChart, kind: 'bogus' }],
  // @ts-expect-error Exercise the JavaScript boundary.
  ['grouping', { ...baseChart, grouping: 'bogus' }],
  // @ts-expect-error Exercise the JavaScript boundary.
  ['valueAxisMajorTickMark', { ...baseChart, valueAxisMajorTickMark: 'bogus' }],
  // @ts-expect-error Exercise the JavaScript boundary.
  ['valueAxisMinorTickMark', { ...baseChart, valueAxisMinorTickMark: 'bogus' }],
  // @ts-expect-error Exercise the JavaScript boundary.
  ['categoryAxisMajorTickMark', { ...baseChart, categoryAxisMajorTickMark: 'bogus' }],
  // @ts-expect-error Exercise the JavaScript boundary.
  ['categoryAxisMinorTickMark', { ...baseChart, categoryAxisMinorTickMark: 'bogus' }],
  // @ts-expect-error Exercise the JavaScript boundary.
  ['valueAxisTickLabelPos', { ...baseChart, valueAxisTickLabelPos: 'bogus' }],
  // @ts-expect-error Exercise the JavaScript boundary.
  ['categoryAxisTickLabelPos', { ...baseChart, categoryAxisTickLabelPos: 'bogus' }],
  // @ts-expect-error Exercise the JavaScript boundary.
  ['categoryAxisLabelAlign', { ...baseChart, categoryAxisLabelAlign: 'bogus' }],
  // @ts-expect-error Exercise the JavaScript boundary.
  ['categoryAxisOrientation', { ...baseChart, categoryAxisOrientation: 'bogus' }],
  // @ts-expect-error Exercise the JavaScript boundary.
  ['valueAxisOrientation', { ...baseChart, valueAxisOrientation: 'bogus' }],
  // @ts-expect-error Exercise the JavaScript boundary.
  ['valueAxisCrosses', { ...baseChart, valueAxisCrosses: 'bogus' }],
  // @ts-expect-error Exercise the JavaScript boundary.
  ['valueAxisCrossBetween', { ...baseChart, valueAxisCrossBetween: 'bogus' }],
  // @ts-expect-error Exercise the JavaScript boundary.
  ['dispBlanksAs', { ...baseChart, dispBlanksAs: 'bogus' }],
  // @ts-expect-error Exercise the JavaScript boundary.
  ['scatterStyle', { ...baseChart, scatterStyle: 'bogus' }],
  // @ts-expect-error Exercise the JavaScript boundary.
  ['radarStyle', { ...baseChart, radarStyle: 'bogus' }],
  // @ts-expect-error Exercise the JavaScript boundary.
  ['bubbleSizeRepresents', { ...baseChart, bubbleSizeRepresents: 'bogus' }],
  // @ts-expect-error Exercise the JavaScript boundary.
  ['bar3DShape', { ...baseChart, bar3DShape: 'bogus' }],
  // @ts-expect-error Exercise the JavaScript boundary.
  ['valueAxis.displayUnits', { ...baseChart, valueAxis: { displayUnits: 'bogus' } }],
  [
    'categoryAxisScaling.displayUnits',
    // @ts-expect-error Exercise the JavaScript boundary.
    { ...baseChart, categoryAxisScaling: { displayUnits: 'bogus' } },
  ],
  [
    'secondaryValueAxis.majorTickMark',
    // @ts-expect-error Exercise the JavaScript boundary.
    { ...baseChart, secondaryValueAxis: { majorTickMark: 'bogus' } },
  ],
  [
    'secondaryValueAxis.minorTickMark',
    // @ts-expect-error Exercise the JavaScript boundary.
    { ...baseChart, secondaryValueAxis: { minorTickMark: 'bogus' } },
  ],
  [
    'secondaryValueAxis.tickLabelPos',
    // @ts-expect-error Exercise the JavaScript boundary.
    { ...baseChart, secondaryValueAxis: { tickLabelPos: 'bogus' } },
  ],
  [
    'secondaryValueAxis.crossBetween',
    // @ts-expect-error Exercise the JavaScript boundary.
    { ...baseChart, secondaryValueAxis: { crossBetween: 'bogus' } },
  ],
  // @ts-expect-error Exercise the JavaScript boundary.
  ['seriesAxis.orientation', { ...baseChart, seriesAxis: { orientation: 'bogus' } }],
  // @ts-expect-error Exercise the JavaScript boundary.
  ['seriesAxis.tickLabelPos', { ...baseChart, seriesAxis: { tickLabelPos: 'bogus' } }],
  // @ts-expect-error Exercise the JavaScript boundary.
  ['categoryAxisDate.baseTimeUnit', { ...baseChart, categoryAxisDate: { baseTimeUnit: 'bogus' } }],
  [
    'categoryAxisDate.majorTimeUnit',
    // @ts-expect-error Exercise the JavaScript boundary.
    { ...baseChart, categoryAxisDate: { majorTimeUnit: 'bogus' } },
  ],
  [
    'categoryAxisDate.minorTimeUnit',
    // @ts-expect-error Exercise the JavaScript boundary.
    { ...baseChart, categoryAxisDate: { minorTimeUnit: 'bogus' } },
  ],
  // @ts-expect-error Exercise the JavaScript boundary.
  ['ofPie.type', { ...baseChart, ofPie: { type: 'bogus' } }],
  // @ts-expect-error Exercise the JavaScript boundary.
  ['ofPie.splitType', { ...baseChart, ofPie: { type: 'pie', splitType: 'bogus' } }],
  [
    'plotAreaLayout.target',
    // @ts-expect-error Exercise the JavaScript boundary.
    { ...baseChart, plotAreaLayout: { x: 0, y: 0, w: 1, h: 1, target: 'bogus' } },
  ],
  [
    'secondaryValueAxis.scaling.displayUnits',
    // @ts-expect-error Exercise the JavaScript boundary.
    { ...baseChart, secondaryValueAxis: { scaling: { displayUnits: 'bogus' } } },
  ],
  // @ts-expect-error Exercise the JavaScript boundary.
  ['legend.position', { ...baseChart, legend: { position: 'bogus' } }],
  [
    'legend.layout.target',
    {
      ...baseChart,
      // @ts-expect-error Exercise the JavaScript boundary.
      legend: { position: 'r', layout: { x: 0, y: 0, w: 1, h: 1, target: 'bogus' } },
    },
  ],
  // @ts-expect-error Exercise the JavaScript boundary.
  ['dataLabels.position', { ...baseChart, dataLabels: { ...labels, position: 'bogus' } }],
  // @ts-expect-error Exercise the JavaScript boundary.
  ['series.chartKind', { ...baseChart, series: [{ ...series, chartKind: 'bogus' }] }],
  ['series.lineDash', { ...baseChart, series: [{ ...series, lineDash: 'bogus' }] }],
  // @ts-expect-error Exercise the JavaScript boundary.
  ['series.markerSymbol', { ...baseChart, series: [{ ...series, markerSymbol: 'bogus' }] }],
  // @ts-expect-error Exercise the JavaScript boundary.
  ['series.trendline', { ...baseChart, series: [{ ...series, trendline: { type: 'bogus' } }] }],
  [
    'series.dataLabels',
    // @ts-expect-error Exercise the JavaScript boundary.
    { ...baseChart, series: [{ ...series, dataLabels: { ...labels, position: 'bogus' } }] },
  ],
  [
    'series.pointDataLabels',
    // @ts-expect-error Exercise the JavaScript boundary.
    { ...baseChart, series: [{ ...series, pointDataLabels: [{ ...labels, position: 'bogus' }] }] },
  ],
  [
    'series.errorBars.barType',
    {
      ...baseChart,
      // @ts-expect-error Exercise the JavaScript boundary.
      series: [{ ...series, errorBars: { barType: 'bogus', amount: { type: 'stdErr' } } }],
    },
  ],
  [
    'series.errorBars.amount',
    {
      ...baseChart,
      // @ts-expect-error Exercise the JavaScript boundary.
      series: [{ ...series, errorBars: { barType: 'both', amount: { type: 'bogus' } } }],
    },
  ],
  [
    'series.xErrorBars.barType',
    {
      ...baseChart,
      // @ts-expect-error Exercise the JavaScript boundary.
      series: [{ ...series, xErrorBars: { barType: 'bogus', amount: { type: 'stdErr' } } }],
    },
  ],
  [
    'series.xErrorBars.amount',
    {
      ...baseChart,
      // @ts-expect-error Exercise the JavaScript boundary.
      series: [{ ...series, xErrorBars: { barType: 'both', amount: { type: 'bogus' } } }],
    },
  ],
];

it.each(invalidCharts)(
  'rejects %s on chart add and update before changing the package',
  async (_, spec) => {
    const pres = api.createPresentation();
    const slide = api.addBlankSlide(pres);
    api.addSlideChart(slide, { ...box, spec: baseChart });
    const chart = api.getSlideCharts(slide)[0]!;
    const before = unzipSync(await api.savePresentation(pres));
    expect(() => api.addSlideChart(slide, { ...box, spec })).toThrow(
      /addSlideChart: .*is not one of:/,
    );
    expect(() => api.setChartSpec(chart, spec)).toThrow(/setChartSpec: .*is not one of:/);
    // ZIP timestamps can change between saves even when every part is unchanged.
    expect(unzipSync(await api.savePresentation(pres))).toEqual(before);
  },
);

it('preserves known enum values on chart add and update across save/load', async () => {
  const spec: api.ChartSpec = {
    ...baseChart,
    grouping: 'stacked',
    valueAxisMajorTickMark: 'out',
    valueAxisMinorTickMark: 'cross',
    categoryAxisMajorTickMark: 'in',
    categoryAxisMinorTickMark: 'none',
    categoryAxisTickLabelPos: 'low',
    valueAxisTickLabelPos: 'high',
    categoryAxisLabelAlign: 'r',
    categoryAxisOrientation: 'maxMin',
    valueAxisOrientation: 'minMax',
    valueAxisCrosses: 'max',
    valueAxisCrossBetween: 'midCat',
    dispBlanksAs: 'span',
    valueAxis: { displayUnits: 'thousands' },
    legend: { position: 'b' },
    dataLabels: { ...labels, position: 'outEnd' },
  };
  const pres = api.createPresentation();
  const slide = api.addBlankSlide(pres);
  api.addSlideChart(slide, { ...box, spec });
  const xmlBefore = new TextDecoder().decode(api.readPackagePart(pres, '/ppt/charts/chart1.xml')!);
  api.setChartSpec(api.getSlideCharts(slide)[0]!, spec);
  const reloaded = await api.loadPresentation(await api.savePresentation(pres));
  const xml = new TextDecoder().decode(api.readPackagePart(reloaded, '/ppt/charts/chart1.xml')!);
  expect(xml).toBe(xmlBefore);
  for (const token of [
    'grouping val="stacked"',
    'minorTickMark val="cross"',
    'majorTickMark val="in"',
    'tickLblPos val="low"',
    'tickLblPos val="high"',
    'lblAlgn val="r"',
    'orientation val="maxMin"',
    'crosses val="max"',
    'crossBetween val="midCat"',
    'dispBlanksAs val="span"',
    'builtInUnit val="thousands"',
    'legendPos val="b"',
    'dLblPos val="outEnd"',
  ])
    expect(xml).toContain(token);
  if (isSchemaValidationAvailable()) expectSchemaValid(xml, 'chart');
});
