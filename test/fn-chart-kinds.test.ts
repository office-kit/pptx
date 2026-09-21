// Every plot-group element of CT_PlotArea is authorable.
//
// For each chart kind and modifier: author it, check the emitted
// `<c:chartSpace>` against the ECMA-376 chart XSD, and read it back through
// `getShapeChartSpec` — the round trip the library promises for everything
// it claims to support.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  type ChartSpec,
  addSlideChart,
  getShapeChartSpec,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  readPackagePart,
  savePresentation,
  setChartSpec,
  getSlideCharts,
} from '../src/api/index.ts';
import { readZip } from '../src/internal/opc/zip.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const decode = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);

interface Authored {
  readonly xml: string;
  readonly sheetXml: string;
  readonly readBack: ChartSpec;
}

// Writes `spec` into a fresh deck and returns the saved deck.
const writeChart = async (spec: ChartSpec): Promise<Uint8Array> => {
  const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
  addSlideChart(getSlides(pres)[0]!, {
    x: inches(0.5),
    y: inches(0.5),
    w: inches(6),
    h: inches(4),
    spec,
  });
  return savePresentation(pres);
};

// Authors `spec`, saves, reloads, and returns the chart XML, the embedded
// sheet XML, and the spec read back from the reloaded deck.
const author = async (spec: ChartSpec): Promise<Authored> => {
  const reloaded = await loadPresentation(await writeChart(spec));
  const xml = decode(readPackagePart(reloaded, '/ppt/charts/chart1.xml')!);
  const xlsx = readZip(
    readPackagePart(reloaded, '/ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx')!,
  );
  const sheetXml = decode(xlsx.entries.find((e) => e.name === 'xl/worksheets/sheet1.xml')!.data);
  const chartShape = getSlideShapes(getSlides(reloaded)[0]!).at(-1)!;
  const readBack = getShapeChartSpec(chartShape);
  if (readBack === null) throw new Error('chart did not read back');
  if (isSchemaValidationAvailable()) expectSchemaValid(xml, 'chart');
  // Round trip: what was read back must write the very same chart. Anything
  // the reader drops or the builder defaults differently shows up here.
  const rewritten = await loadPresentation(await writeChart(readBack));
  expect(decode(readPackagePart(rewritten, '/ppt/charts/chart1.xml')!)).toBe(xml);
  return { xml, sheetXml, readBack };
};

const CATEGORIES = ['Q1', 'Q2', 'Q3', 'Q4'];
const TWO_SERIES = [
  { name: 'North', values: [10, 20, 15, 30] },
  { name: 'South', values: [12, 18, 22, 25] },
];

describe('chart kinds: xy charts', () => {
  it('scatter chart writes per-series x / y channels and reads them back', async () => {
    const { xml, sheetXml, readBack } = await author({
      kind: 'scatter',
      categories: [],
      series: [
        { name: 'Trial A', xValues: [1, 2, 3], values: [2.5, 4.1, 6.2] },
        { name: 'Trial B', xValues: [1.5, 2.5], values: [3, 5] },
      ],
      scatterStyle: 'lineMarker',
      categoryAxisTitle: 'Dose',
      categoryAxisScaling: { min: 0, max: 4, majorUnit: 1 },
      valueAxisTitle: 'Response',
      valueAxis: { min: 0, max: 10 },
    });
    expect(xml).toContain('<c:scatterChart>');
    expect(xml).toContain('<c:xVal>');
    expect(xml).toContain('<c:yVal>');
    // Each series owns a block of columns: A/B for the first, C/D for the second.
    expect(xml).toContain('Sheet1!$A$2:$A$4');
    expect(xml).toContain('Sheet1!$B$2:$B$4');
    expect(xml).toContain('Sheet1!$C$2:$C$3');
    expect(xml).toContain('Sheet1!$D$2:$D$3');
    expect(sheetXml).toContain('<c r="C2"><v>1.5</v></c>');
    // Both axes of a scatter chart are value axes.
    expect(xml.match(/<c:valAx>/g)).toHaveLength(2);
    expect(xml).not.toContain('<c:catAx>');

    expect(readBack.kind).toBe('scatter');
    expect(readBack.scatterStyle).toBe('lineMarker');
    expect(readBack.series.map((s) => s.xValues)).toEqual([
      [1, 2, 3],
      [1.5, 2.5],
    ]);
    expect(readBack.series.map((s) => s.values)).toEqual([
      [2.5, 4.1, 6.2],
      [3, 5],
    ]);
    expect(readBack.categoryAxisTitle).toBe('Dose');
    expect(readBack.categoryAxisScaling).toEqual({ min: 0, max: 4, majorUnit: 1 });
    expect(readBack.valueAxisTitle).toBe('Response');
    expect(readBack.valueAxis).toMatchObject({ min: 0, max: 10 });
  });

  it('markers-only scatter hides the connecting line on every series', async () => {
    const { xml } = await author({
      kind: 'scatter',
      categories: [],
      series: [{ name: 'Points', xValues: [1, 2], values: [3, 4] }],
    });
    expect(xml).toContain('<a:ln><a:noFill/></a:ln>');
    expect(xml).toContain('<c:smooth val="0"/>');
  });

  it('bubble chart writes sizes, scale and the 3-D flag', async () => {
    const { xml, readBack } = await author({
      kind: 'bubble',
      categories: [],
      series: [
        {
          name: 'Markets',
          xValues: [10, 20, 30],
          values: [5, 9, 4],
          bubbleSizes: [100, 250, 60],
          fillOpacity: 0.6,
        },
      ],
      bubbleScale: 80,
      bubbleSizeRepresents: 'width',
      bubble3D: true,
      showNegativeBubbles: false,
      dataLabels: {
        showValue: false,
        showCategory: false,
        showSeriesName: false,
        showPercent: false,
        showBubbleSize: true,
      },
    });
    expect(xml).toContain('<c:bubbleChart>');
    expect(xml).toContain('<c:bubbleSize>');
    expect(xml).toContain('<c:bubbleScale val="80"/>');
    expect(xml).toContain('<c:sizeRepresents val="w"/>');
    expect(xml).toContain('<a:alpha val="60000"/>');
    // PowerPoint repairs a deck with `<c:bubble3D>` directly under
    // `<c:bubbleChart>`; the flag belongs to the series alone.
    expect(xml).toContain('<c:bubble3D val="1"/></c:ser>');
    expect(xml.split('<c:bubble3D').length - 1).toBe(1);
    expect(readBack.kind).toBe('bubble');
    expect(readBack.series[0]!.bubbleSizes).toEqual([100, 250, 60]);
    expect(readBack.series[0]!.fillOpacity).toBe(0.6);
    expect(readBack.bubbleScale).toBe(80);
    expect(readBack.bubbleSizeRepresents).toBe('width');
    expect(readBack.bubble3D).toBe(true);
    expect(readBack.dataLabels?.showBubbleSize).toBe(true);
  });

  it('rejects an xy series without its x channel', async () => {
    await expect(
      author({ kind: 'scatter', categories: [], series: [{ name: 'S', values: [1, 2] }] }),
    ).rejects.toThrow(/needs xValues/);
    await expect(
      author({
        kind: 'bubble',
        categories: [],
        series: [{ name: 'S', xValues: [1], values: [1] }],
      }),
    ).rejects.toThrow(/needs bubbleSizes/);
  });
});

describe('chart kinds: radar, stock, surface', () => {
  it('radar chart round-trips its style', async () => {
    const { xml, readBack } = await author({
      kind: 'radar',
      categories: ['Speed', 'Range', 'Comfort', 'Price', 'Safety'],
      series: [
        { name: 'Model X', values: [4, 3, 5, 2, 5] },
        { name: 'Model Y', values: [3, 5, 3, 4, 4], fillOpacity: 0.4 },
      ],
      radarStyle: 'filled',
    });
    expect(xml).toContain('<c:radarChart>');
    expect(xml).toContain('<c:radarStyle val="filled"/>');
    expect(readBack.kind).toBe('radar');
    expect(readBack.radarStyle).toBe('filled');
    expect(readBack.series[1]!.values).toEqual([3, 5, 3, 4, 4]);
  });

  it('stock chart (OHLC) carries high-low lines and up/down bars', async () => {
    const { xml, readBack } = await author({
      kind: 'stock',
      categories: ['Mon', 'Tue', 'Wed'],
      series: [
        { name: 'Open', values: [10, 12, 11] },
        { name: 'High', values: [14, 15, 13] },
        { name: 'Low', values: [9, 11, 8] },
        { name: 'Close', values: [12, 11, 12] },
      ],
      upDownBars: { upColor: '#2E7D32', downColor: '#C62828', gapWidthPct: 100 },
    });
    expect(xml).toContain('<c:stockChart>');
    expect(xml).toContain('<c:hiLowLines/>');
    expect(xml).toContain('<c:upDownBars>');
    expect(readBack.kind).toBe('stock');
    expect(readBack.series.map((s) => s.name)).toEqual(['Open', 'High', 'Low', 'Close']);
    expect(readBack.hiLowLines).toBe(true);
    expect(readBack.upDownBars).toEqual({
      upColor: '#2E7D32',
      downColor: '#C62828',
      gapWidthPct: 100,
    });
  });

  it('stock chart rejects a series count CT_StockChart cannot hold', async () => {
    await expect(
      author({ kind: 'stock', categories: ['Mon'], series: TWO_SERIES }),
    ).rejects.toThrow(/3 series .* or 4/);
  });

  it('surface chart writes three axes and a camera', async () => {
    const { xml, readBack } = await author({
      kind: 'surface',
      categories: ['0', '10', '20'],
      series: [
        { name: 'Row 1', values: [1, 2, 3] },
        { name: 'Row 2', values: [2, 4, 6] },
        { name: 'Row 3', values: [3, 6, 9] },
      ],
      surfaceWireframe: true,
      seriesAxis: { title: 'Depth' },
    });
    expect(xml).toContain('<c:surface3DChart>');
    expect(xml).toContain('<c:wireframe val="1"/>');
    expect(xml).toContain('<c:serAx>');
    expect(xml).toContain('<c:view3D>');
    expect(readBack.kind).toBe('surface');
    expect(readBack.surfaceWireframe).toBe(true);
    expect(readBack.surfaceContour).toBeUndefined();
    expect(readBack.seriesAxis?.title).toBe('Depth');
  });

  it('contour surface is the top-down surfaceChart', async () => {
    const { xml, readBack } = await author({
      kind: 'surface',
      categories: ['0', '10'],
      series: [
        { name: 'Row 1', values: [1, 2] },
        { name: 'Row 2', values: [2, 4] },
      ],
      surfaceContour: true,
    });
    expect(xml).toContain('<c:surfaceChart>');
    expect(xml).toContain('<c:rotX val="90"/>');
    expect(readBack.surfaceContour).toBe(true);
  });
});

describe('chart kinds: 3-D variants', () => {
  it('3-D column chart with a bar shape', async () => {
    const { xml, readBack } = await author({
      kind: 'column',
      categories: CATEGORIES,
      series: TWO_SERIES,
      view3D: { rotX: 20, rotY: 30, rightAngleAxes: false, perspective: 40 },
      bar3DShape: 'cylinder',
      gapDepthPct: 120,
    });
    expect(xml).toContain('<c:bar3DChart>');
    expect(xml).toContain('<c:shape val="cylinder"/>');
    expect(xml).toContain('<c:gapDepth val="120"/>');
    expect(xml).not.toContain('<c:overlap');
    expect(readBack.kind).toBe('column');
    expect(readBack.view3D).toEqual({
      rotX: 20,
      rotY: 30,
      rightAngleAxes: false,
      perspective: 40,
    });
    expect(readBack.bar3DShape).toBe('cylinder');
    expect(readBack.gapDepthPct).toBe(120);
  });

  it('3-D bar chart with series laid out in depth gets a series axis', async () => {
    const { xml } = await author({
      kind: 'bar',
      categories: CATEGORIES,
      series: TWO_SERIES,
      grouping: 'standard',
      view3D: {},
    });
    expect(xml).toContain('<c:barDir val="bar"/>');
    expect(xml).toContain('<c:serAx>');
    expect(xml.match(/<c:axId val=/g)!.length).toBeGreaterThanOrEqual(6);
  });

  it('3-D line, area and pie select their 3-D elements', async () => {
    const line = await author({
      kind: 'line',
      categories: CATEGORIES,
      series: TWO_SERIES,
      view3D: {},
    });
    expect(line.xml).toContain('<c:line3DChart>');
    expect(line.xml).toContain('<c:serAx>');
    expect(line.readBack.view3D).toBeDefined();

    const area = await author({
      kind: 'area',
      categories: CATEGORIES,
      series: TWO_SERIES,
      grouping: 'stacked',
      view3D: { depthPercent: 200 },
    });
    expect(area.xml).toContain('<c:area3DChart>');
    expect(area.readBack.view3D?.depthPercent).toBe(200);

    const pie = await author({
      kind: 'pie',
      categories: CATEGORIES,
      series: [TWO_SERIES[0]!],
      view3D: { rotX: 45 },
    });
    expect(pie.xml).toContain('<c:pie3DChart>');
    expect(pie.readBack.kind).toBe('pie');
    expect(pie.readBack.view3D?.rotX).toBe(45);
  });

  it('rejects view3D where no 3-D element exists', async () => {
    await expect(
      author({ kind: 'doughnut', categories: CATEGORIES, series: [TWO_SERIES[0]!], view3D: {} }),
    ).rejects.toThrow(/no 3-D variant/);
    await expect(
      author({
        kind: 'pie',
        categories: CATEGORIES,
        series: [TWO_SERIES[0]!],
        view3D: {},
        firstSliceAngleDeg: 90,
      }),
    ).rejects.toThrow(/view3D\.rotY/);
  });
});

describe('chart kinds: pie of pie', () => {
  it('bar-of-pie with a custom split', async () => {
    const ofPie = {
      type: 'bar',
      splitType: 'cust',
      customSplit: [2, 3],
      secondPieSizePct: 60,
      gapWidthPct: 80,
      seriesLines: true,
    } as const;
    const { xml, readBack } = await author({
      kind: 'pie',
      categories: CATEGORIES,
      series: [TWO_SERIES[0]!],
      ofPie,
    });
    expect(xml).toContain('<c:ofPieChart>');
    expect(xml).toContain('<c:ofPieType val="bar"/>');
    expect(xml).toContain('<c:secondPiePt val="3"/>');
    expect(readBack.kind).toBe('pie');
    expect(readBack.ofPie).toEqual(ofPie);
  });
});

describe('chart features shared across kinds', () => {
  it('error bars: fixed on a column series, custom x / y on a scatter series', async () => {
    const column = await author({
      kind: 'column',
      categories: CATEGORIES,
      series: [
        {
          ...TWO_SERIES[0]!,
          errorBars: {
            barType: 'both',
            amount: { type: 'percentage', value: 10 },
            noEndCap: true,
            color: '#333333',
            lineWidthEmu: 12700,
          },
        },
      ],
    });
    expect(column.xml).toContain('<c:errValType val="percentage"/>');
    expect(column.readBack.series[0]!.errorBars).toEqual({
      barType: 'both',
      amount: { type: 'percentage', value: 10 },
      noEndCap: true,
      color: '#333333',
      lineWidthEmu: 12700,
    });

    const scatter = await author({
      kind: 'scatter',
      categories: [],
      series: [
        {
          name: 'S',
          xValues: [1, 2],
          values: [3, 4],
          errorBars: { barType: 'plus', amount: { type: 'cust', plus: [0.5, 0.25] } },
          xErrorBars: { barType: 'both', amount: { type: 'stdErr' } },
        },
      ],
    });
    expect(scatter.xml).toContain('<c:errDir val="x"/>');
    expect(scatter.xml).toContain('<c:errDir val="y"/>');
    expect(scatter.readBack.series[0]!.errorBars).toEqual({
      barType: 'plus',
      amount: { type: 'cust', plus: [0.5, 0.25] },
    });
    expect(scatter.readBack.series[0]!.xErrorBars).toEqual({
      barType: 'both',
      amount: { type: 'stdErr' },
    });
  });

  it('rejects error bars on a series type that has no errBars child', async () => {
    await expect(
      author({
        kind: 'pie',
        categories: CATEGORIES,
        series: [{ ...TWO_SERIES[0]!, errorBars: { barType: 'both', amount: { type: 'stdErr' } } }],
      }),
    ).rejects.toThrow(/error bars/);
  });

  it('data table, plot-area / legend / title layout', async () => {
    const { xml, readBack } = await author({
      kind: 'column',
      categories: CATEGORIES,
      series: TWO_SERIES,
      title: 'Placed',
      titleLayout: { x: 0.05, y: 0.02 },
      dataTable: { showKeys: false, textStyle: { sizePt: 9 } },
      plotAreaLayout: { x: 0.1, y: 0.15, w: 0.8, h: 0.6, target: 'inner' },
      legend: { position: 'r', layout: { x: 0.85, y: 0.4, w: 0.12, h: 0.2 } },
    });
    expect(xml).toContain('<c:dTable>');
    expect(xml).toContain('<c:layoutTarget val="inner"/>');
    expect(readBack.dataTable).toEqual({
      showHorizontalBorder: true,
      showVerticalBorder: true,
      showOutline: true,
      showKeys: false,
      textStyle: { sizePt: 9 },
    });
    expect(readBack.plotAreaLayout).toEqual({ x: 0.1, y: 0.15, w: 0.8, h: 0.6, target: 'inner' });
    expect(readBack.legend?.layout).toEqual({ x: 0.85, y: 0.4, w: 0.12, h: 0.2 });
    expect(readBack.titleLayout).toEqual({ x: 0.05, y: 0.02 });
  });

  it('date axis writes numeric categories and time units', async () => {
    const { xml, sheetXml, readBack } = await author({
      kind: 'line',
      categories: ['45292', '45323', '45352'],
      series: [{ name: 'Users', values: [100, 140, 180] }],
      categoryAxisDate: { baseTimeUnit: 'months', majorUnit: 1, majorTimeUnit: 'months' },
      categoryAxisNumberFormat: 'yyyy-mm',
      categoryAxisScaling: { min: 45292, max: 45352 },
    });
    expect(xml).toContain('<c:dateAx>');
    expect(xml).not.toContain('<c:catAx>');
    expect(xml).toContain('<c:baseTimeUnit val="months"/>');
    // Date categories are numbers in the sheet, not text, so Excel can space them.
    expect(sheetXml).toContain('<c r="A2"><v>45292</v></c>');
    expect(readBack.categories).toEqual(['45292', '45323', '45352']);
    expect(readBack.categoryAxisDate).toEqual({
      baseTimeUnit: 'months',
      majorUnit: 1,
      majorTimeUnit: 'months',
    });
    expect(readBack.categoryAxisScaling).toEqual({ min: 45292, max: 45352 });
  });

  it('date axis rejects non-numeric categories', async () => {
    await expect(
      author({
        kind: 'line',
        categories: ['Jan', 'Feb'],
        series: [{ name: 'S', values: [1, 2] }],
        categoryAxisDate: {},
      }),
    ).rejects.toThrow(/date serial/);
  });

  it('multi-level categories span every level column', async () => {
    const levels = [['2024', '', '2025', '']];
    const { xml, sheetXml, readBack } = await author({
      kind: 'column',
      categories: ['H1', 'H2', 'H1', 'H2'],
      categoryGroupLevels: levels,
      series: [{ name: 'Revenue', values: [1, 2, 3, 4] }],
    });
    expect(xml).toContain('<c:multiLvlStrRef>');
    expect(xml).toContain('Sheet1!$A$2:$B$5');
    expect(xml).toContain('Sheet1!$C$2:$C$5');
    // Outer level in column A, the categories themselves in column B.
    expect(sheetXml).toContain('<c r="A2" t="inlineStr"><is><t>2024</t></is></c>');
    expect(sheetXml).toContain('<c r="B2" t="inlineStr"><is><t>H1</t></is></c>');
    expect(readBack.categories).toEqual(['H1', 'H2', 'H1', 'H2']);
    expect(readBack.categoryGroupLevels).toEqual(levels);
  });

  it('per-point label text, label fill, axis-line hiding, display-unit caption', async () => {
    const labels = {
      showValue: true,
      showCategory: false,
      showSeriesName: false,
      showPercent: false,
    };
    const { xml, readBack } = await author({
      kind: 'column',
      categories: CATEGORIES,
      series: [
        {
          ...TWO_SERIES[0]!,
          dataLabels: { ...labels, fillColor: '#FFF2CC' },
          pointDataLabels: [null, { ...labels, text: 'Peak' }],
        },
      ],
      valueAxis: { displayUnits: 'thousands', displayUnitsLabel: true },
      valueAxisLineHidden: true,
      categoryAxisLineHidden: true,
    });
    expect(xml).toContain('<c:dispUnitsLbl/>');
    expect(readBack.series[0]!.dataLabels?.fillColor).toBe('#FFF2CC');
    expect(readBack.series[0]!.pointDataLabels?.[1]?.text).toBe('Peak');
    expect(readBack.valueAxis).toMatchObject({
      displayUnits: 'thousands',
      displayUnitsLabel: true,
    });
    expect(readBack.valueAxisLineHidden).toBe(true);
    expect(readBack.categoryAxisLineHidden).toBe(true);
  });

  it('formulas stay correct past column Z', async () => {
    const series = Array.from({ length: 30 }, (_, i) => ({ name: `S${i}`, values: [i, i + 1] }));
    const { xml } = await author({ kind: 'line', categories: ['a', 'b'], series });
    // Series 25 is the 27th column: AA, not the `[` that charCode math yields.
    expect(xml).toContain('Sheet1!$AA$2:$AA$3');
    expect(xml).not.toContain('$[');
  });

  it('setChartSpec switches an existing chart to an xy kind, workbook included', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: { kind: 'column', categories: CATEGORIES, series: TWO_SERIES },
    });
    setChartSpec(getSlideCharts(slide)[0]!, {
      kind: 'scatter',
      categories: [],
      series: [{ name: 'XY', xValues: [1, 2], values: [3, 4] }],
    });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const chart = getSlideCharts(getSlides(reloaded)[0]!)[0]!;
    expect(chart.spec?.kind).toBe('scatter');
    const xlsx = readZip(
      readPackagePart(reloaded, '/ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx')!,
    );
    const sheetXml = decode(xlsx.entries.find((e) => e.name === 'xl/worksheets/sheet1.xml')!.data);
    expect(sheetXml).toContain('<c r="B1" t="inlineStr"><is><t>XY</t></is></c>');
  });
});
