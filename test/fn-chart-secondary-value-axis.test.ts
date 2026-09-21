// `ChartSpec.secondaryValueAxis` — formatting for the right-hand axis that
// `secondaryAxis` series plot against — plus the axis-position and
// point-count fixes that landed with it.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';
import { readChartSpec } from '../src/internal/chartml/index.ts';
import { parseXml } from '../src/internal/xml/index.ts';
import {
  type ChartSecondaryValueAxis,
  type ChartSpec,
  type ReadChartSpec,
  addSlideChart,
  getSlideCharts,
  getSlides,
  inches,
  loadPresentation,
  readPackagePart,
  savePresentation,
  setChartSpec,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const decoder = new TextDecoder();

const roundTrip = async (
  spec: ChartSpec,
): Promise<{ readonly spec: ReadChartSpec; readonly xml: string }> => {
  const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
  addSlideChart(getSlides(pres)[0]!, {
    x: inches(0.5),
    y: inches(0.5),
    w: inches(8),
    h: inches(4.5),
    spec,
  });
  const reloaded = await loadPresentation(await savePresentation(pres));
  const charts = getSlideCharts(getSlides(reloaded)[0]!);
  expect(charts).toHaveLength(1);
  const xml = decoder.decode(readPackagePart(reloaded, '/ppt/charts/chart1.xml')!);
  return { spec: charts[0]!.spec!, xml };
};

const sliceElement = (xml: string, tag: string): string =>
  xml.slice(xml.indexOf(`<c:${tag}>`), xml.indexOf(`</c:${tag}>`));

const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

describe('fn API: secondaryValueAxis', () => {
  const secondary: ChartSecondaryValueAxis = {
    scaling: { min: 0, max: 50, majorUnit: 10, numberFormat: '#,##0' },
    title: 'Score',
    titleStyle: { font: 'Yu Gothic', sizePt: 12, bold: true },
    labelStyle: { font: 'Yu Gothic', sizePt: 12, color: '#000000' },
    majorGridlines: true,
    majorGridlineColor: '#888888',
    lineColor: '#888888',
    majorTickMark: 'out',
    minorTickMark: 'none',
    crossBetween: 'between',
  };

  skipIfNoXmllint(
    'round-trips every secondary-axis field independently of the primary axis',
    async () => {
      const { spec, xml } = await roundTrip({
        kind: 'column',
        categories: ['A', 'B', 'C'],
        series: [
          { name: 'Count', values: [92, 118, 54] },
          { name: 'Score', values: [4.5, 3.8, 2.9], chartKind: 'line', secondaryAxis: true },
        ],
        valueAxis: { min: 0, max: 200 },
        secondaryValueAxis: secondary,
      });
      // The deck body-text color is baked onto the unset title color.
      expect(spec.secondaryValueAxis).toEqual({
        ...secondary,
        titleStyle: { ...secondary.titleStyle, color: '#000000' },
      });
      expect(spec.valueAxis).toEqual({ min: 0, max: 200 });
      // The secondary axis keeps its fixed position and crossing.
      const secondaryXml = xml.slice(xml.indexOf('<c:axId val="444444444"/>'));
      expect(secondaryXml).toContain('<c:axPos val="r"/>');
      expect(secondaryXml).toContain('<c:crosses val="max"/>');
      expect(secondaryXml).toContain('<c:majorUnit val="10"/>');
      expect(secondaryXml).toContain('<c:minorTickMark val="none"/>');
      expectSchemaValid(xml, 'chart');
    },
  );

  it('bakes the deck text color on the secondary axis labels when none was authored', async () => {
    const { spec } = await roundTrip({
      kind: 'column',
      categories: ['A', 'B'],
      series: [
        { name: 'Count', values: [1, 2] },
        { name: 'Score', values: [3, 4], chartKind: 'line', secondaryAxis: true },
      ],
    });
    expect(spec.secondaryValueAxis).toEqual({ labelStyle: { color: '#000000' } });
    expect(spec.series[1]!.secondaryAxis).toBe(true);
  });

  it('rejects a chart whose every series plots on the secondary axis', async () => {
    await expect(
      roundTrip({
        kind: 'column',
        categories: ['A', 'B'],
        series: [{ name: 'only', values: [1, 2], secondaryAxis: true }],
        secondaryValueAxis: { title: 'Right', scaling: { max: 5 } },
      }),
    ).rejects.toThrow(/^chart: at least one series must plot on the primary axis/);
  });

  it('rejects an all-secondary spec on setChartSpec under the same message', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(8),
      h: inches(4.5),
      spec: { kind: 'column', categories: ['A'], series: [{ name: 'a', values: [1] }] },
    });
    expect(() =>
      setChartSpec(getSlideCharts(slide)[0]!, {
        kind: 'column',
        categories: ['A'],
        series: [{ name: 'only', values: [1], secondaryAxis: true }],
      }),
    ).toThrow(/^chart: at least one series must plot on the primary axis/);
  });

  skipIfNoXmllint(
    'round-trips tickLabelPos on the secondary axis of a horizontal bar',
    async () => {
      const { spec, xml } = await roundTrip({
        kind: 'bar',
        categories: ['A', 'B'],
        series: [
          { name: 'Count', values: [1, 2] },
          { name: 'Share', values: [3, 4], secondaryAxis: true },
        ],
        valueAxisTickLabelPos: 'low',
        secondaryValueAxis: { tickLabelPos: 'low', minorTickMark: 'none' },
      });
      const secondaryXml = xml.slice(xml.indexOf('<c:axId val="444444444"/>'));
      expect(secondaryXml).toContain('<c:minorTickMark val="none"/><c:tickLblPos val="low"/>');
      expect(spec.secondaryValueAxis).toMatchObject({ tickLabelPos: 'low', minorTickMark: 'none' });
      expect(spec.valueAxisTickLabelPos).toBe('low');
      expectSchemaValid(xml, 'chart');
    },
  );

  it('is ignored for charts without secondary-axis series', async () => {
    const { spec, xml } = await roundTrip({
      kind: 'column',
      categories: ['A', 'B'],
      series: [{ name: 'Count', values: [1, 2] }],
      secondaryValueAxis: secondary,
    });
    expect(spec.secondaryValueAxis).toBeUndefined();
    expect(xml).not.toContain('<c:axId val="444444444"/>');
  });
});

describe('chart reader: secondary value axis selection', () => {
  it('picks the axis the secondary group references even when it precedes the primary axis', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart><c:plotArea><c:layout/>
    <c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
      <c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:tx>
        <c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>
      <c:axId val="1"/><c:axId val="2"/>
    </c:barChart>
    <c:lineChart><c:grouping val="standard"/>
      <c:ser><c:idx val="1"/><c:order val="1"/><c:tx><c:strLit><c:pt idx="0"><c:v>B</c:v></c:pt></c:strLit></c:tx>
        <c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
      <c:axId val="3"/><c:axId val="4"/>
    </c:lineChart>
    <c:valAx><c:axId val="4"/><c:scaling><c:orientation val="minMax"/><c:max val="9"/></c:scaling><c:delete val="0"/><c:axPos val="r"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:crossAx val="3"/><c:crosses val="max"/></c:valAx>
    <c:catAx><c:axId val="3"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="b"/><c:crossAx val="4"/></c:catAx>
    <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
    <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/><c:max val="5"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/><c:crosses val="autoZero"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const spec = readChartSpec(parseXml(xml).root)!;
    expect(spec.secondaryValueAxis).toEqual({
      scaling: { max: 9 },
      majorTickMark: 'out',
      minorTickMark: 'none',
    });
    expect(spec.valueAxis).toEqual({ max: 5 });
    expect(spec.series[1]!.secondaryAxis).toBe(true);
  });

  // A scatter group references two value axes (X and Y); neither is
  // secondary, whichever side X sits on.
  const scatterWithXAt = (
    xPos: 'b' | 't',
  ): string => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart><c:plotArea><c:layout/>
    <c:scatterChart><c:scatterStyle val="lineMarker"/>
      <c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:tx>
        <c:xVal><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:xVal>
        <c:yVal><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:yVal></c:ser>
      <c:axId val="1"/><c:axId val="2"/>
    </c:scatterChart>
    <c:valAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/><c:max val="20"/></c:scaling><c:delete val="0"/><c:axPos val="${xPos}"/><c:crossAx val="2"/><c:crosses val="autoZero"/></c:valAx>
    <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/><c:max val="40"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/><c:crosses val="autoZero"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;

  it.each(['b', 't'] as const)(
    'reads a scatter chart with the X axis at %s as primary only',
    (xPos) => {
      const spec = readChartSpec(parseXml(scatterWithXAt(xPos)).root)!;
      expect(spec.secondaryValueAxis).toBeUndefined();
      expect(spec.series[0]!.secondaryAxis).toBeUndefined();
      // The group's first axId is the x axis, wherever it is drawn: its
      // scaling is the horizontal axis', the other one the value axis'.
      expect(spec.categoryAxisScaling).toEqual({ max: 20 });
      expect(spec.valueAxis).toEqual({ max: 40 });
    },
  );
});

describe('fn API: axis line and gridline widths', () => {
  skipIfNoXmllint('round-trips widths on every axis line and gridline', async () => {
    const { spec, xml } = await roundTrip({
      kind: 'column',
      categories: ['A', 'B'],
      series: [
        { name: 'Count', values: [1, 2] },
        { name: 'Share', values: [3, 4], chartKind: 'line', secondaryAxis: true },
      ],
      valueAxisMajorGridlines: true,
      valueAxisMajorGridlineColor: '#888888',
      valueAxisMajorGridlineWidthEmu: 12700,
      valueAxisMinorGridlines: true,
      valueAxisMinorGridlineWidthEmu: 6350,
      categoryAxisMajorGridlines: true,
      categoryAxisMajorGridlineWidthEmu: 19050,
      categoryAxisMinorGridlines: true,
      categoryAxisMinorGridlineColor: '#CCCCCC',
      categoryAxisMinorGridlineWidthEmu: 3175,
      valueAxisLineColor: '#888888',
      valueAxisLineWidthEmu: 12700,
      categoryAxisLineWidthEmu: 25400,
      secondaryValueAxis: {
        majorGridlines: true,
        majorGridlineWidthEmu: 9525,
        lineColor: '#444444',
        lineWidthEmu: 15875,
      },
    });
    const catAx = sliceElement(xml, 'catAx');
    const valAx = sliceElement(xml, 'valAx');
    const secondary = xml.slice(xml.lastIndexOf('<c:valAx>'));
    const colored = (w: number, hex: string): string =>
      `<c:spPr><a:ln w="${w}"><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></a:ln></c:spPr>`;
    expect(valAx).toContain(`<c:majorGridlines>${colored(12700, '888888')}</c:majorGridlines>`);
    // A width without a color is a bare <a:ln w>.
    expect(valAx).toContain(
      '<c:minorGridlines><c:spPr><a:ln w="6350"/></c:spPr></c:minorGridlines>',
    );
    expect(valAx).toContain(colored(12700, '888888') + '<c:txPr>');
    expect(catAx).toContain(
      '<c:majorGridlines><c:spPr><a:ln w="19050"/></c:spPr></c:majorGridlines>',
    );
    expect(catAx).toContain(`<c:minorGridlines>${colored(3175, 'CCCCCC')}</c:minorGridlines>`);
    expect(catAx).toContain('<c:spPr><a:ln w="25400"/></c:spPr><c:txPr>');
    expect(secondary).toContain(
      '<c:majorGridlines><c:spPr><a:ln w="9525"/></c:spPr></c:majorGridlines>',
    );
    expect(secondary).toContain(colored(15875, '444444'));
    expect(spec).toMatchObject({
      valueAxisMajorGridlineWidthEmu: 12700,
      valueAxisMinorGridlineWidthEmu: 6350,
      categoryAxisMajorGridlineWidthEmu: 19050,
      categoryAxisMinorGridlineWidthEmu: 3175,
      valueAxisLineWidthEmu: 12700,
      categoryAxisLineWidthEmu: 25400,
      secondaryValueAxis: {
        majorGridlineWidthEmu: 9525,
        lineColor: '#444444',
        lineWidthEmu: 15875,
      },
    });
    expect(spec.valueAxisMinorGridlineColor).toBeUndefined();
    expect(spec.categoryAxisLineColor).toBeUndefined();
    expectSchemaValid(xml, 'chart');
  });

  skipIfNoXmllint('round-trips a zero width, the lower bound of ST_LineWidth', async () => {
    const { spec, xml } = await roundTrip({
      kind: 'line',
      categories: ['A', 'B'],
      series: [
        { name: 'Trend', values: [1, 2], lineWidthEmu: 0 },
        { name: 'Share', values: [3, 4], secondaryAxis: true },
      ],
      valueAxisMajorGridlines: true,
      valueAxisMajorGridlineWidthEmu: 0,
      valueAxisLineWidthEmu: 0,
      categoryAxisLineColor: '#888888',
      categoryAxisLineWidthEmu: 0,
      secondaryValueAxis: { lineWidthEmu: 0 },
    });
    expect(sliceElement(xml, 'valAx')).toContain(
      '<c:majorGridlines><c:spPr><a:ln w="0"/></c:spPr></c:majorGridlines>',
    );
    expect(sliceElement(xml, 'catAx')).toContain('<c:spPr><a:ln w="0"><a:solidFill>');
    expect(sliceElement(xml, 'ser')).toContain('<a:ln w="0">');
    expect(spec).toMatchObject({
      valueAxisMajorGridlineWidthEmu: 0,
      valueAxisLineWidthEmu: 0,
      categoryAxisLineWidthEmu: 0,
      secondaryValueAxis: { lineWidthEmu: 0 },
    });
    expect(spec.series[0]!.lineWidthEmu).toBe(0);
    expectSchemaValid(xml, 'chart');
  });

  it('writes no width when only the color is authored', async () => {
    const { spec, xml } = await roundTrip({
      kind: 'column',
      categories: ['A', 'B'],
      series: [{ name: 'Count', values: [1, 2] }],
      valueAxisMajorGridlines: true,
      valueAxisMajorGridlineColor: '#888888',
      valueAxisLineColor: '#888888',
    });
    expect(sliceElement(xml, 'valAx')).not.toContain('<a:ln w=');
    expect(spec.valueAxisMajorGridlineWidthEmu).toBeUndefined();
    expect(spec.valueAxisLineWidthEmu).toBeUndefined();
  });

  it('rejects a width outside ST_LineWidth', async () => {
    await expect(
      roundTrip({
        kind: 'column',
        categories: ['A'],
        series: [{ name: 'a', values: [1] }],
        valueAxisLineWidthEmu: -1,
      }),
    ).rejects.toThrow(/valueAxisLineWidthEmu/);
    await expect(
      roundTrip({
        kind: 'column',
        categories: ['A'],
        series: [
          { name: 'a', values: [1] },
          { name: 'b', values: [1], secondaryAxis: true },
        ],
        secondaryValueAxis: { lineWidthEmu: 20116801 },
      }),
    ).rejects.toThrow(/secondaryValueAxis\.lineWidthEmu/);
  });
});

describe('fn API: series markerColor', () => {
  skipIfNoXmllint('round-trips a marker color that differs from the line', async () => {
    const { spec, xml } = await roundTrip({
      kind: 'line',
      categories: ['A', 'B'],
      series: [
        {
          name: 'Trend',
          values: [3, 1],
          color: '#112233',
          markerSymbol: 'circle',
          markerSizePt: 6,
          markerColor: '#AABBCC',
        },
      ],
    });
    // CT_Marker order: symbol, size, spPr. The outline follows the fill.
    expect(sliceElement(xml, 'ser')).toContain(
      '<c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr><a:solidFill><a:srgbClr val="AABBCC"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="AABBCC"/></a:solidFill></a:ln></c:spPr></c:marker>',
    );
    expect(spec.series[0]).toMatchObject({
      color: '#112233',
      markerColor: '#AABBCC',
      markerLineColor: '#AABBCC',
    });
    expectSchemaValid(xml, 'chart');
  });

  skipIfNoXmllint('round-trips a marker outline color separate from its fill', async () => {
    const { spec, xml } = await roundTrip({
      kind: 'line',
      categories: ['A', 'B'],
      series: [
        {
          name: 'Trend',
          values: [3, 1],
          color: '#112233',
          markerSymbol: 'circle',
          markerColor: '#FFFFFF',
          markerLineColor: '#112233',
        },
      ],
    });
    expect(sliceElement(xml, 'ser')).toContain(
      '<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln></c:spPr></c:marker>',
    );
    expect(spec.series[0]).toMatchObject({ markerColor: '#FFFFFF', markerLineColor: '#112233' });
    expectSchemaValid(xml, 'chart');
  });

  skipIfNoXmllint('writes the marker in the series color when nothing is authored', async () => {
    const { spec, xml } = await roundTrip({
      kind: 'line',
      categories: ['A', 'B'],
      series: [{ name: 'Trend', values: [3, 1], color: '#112233' }],
    });
    expect(sliceElement(xml, 'ser')).toContain(
      '<c:marker><c:spPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill>',
    );
    expect(spec.series[0]!.markerColor).toBe('#112233');
    expect(spec.series[0]!.markerLineColor).toBe('#112233');
    expect(spec.series[0]!.markerSymbol).toBeUndefined();
    expectSchemaValid(xml, 'chart');
  });

  skipIfNoXmllint('writes no marker spPr for a hidden marker or a bar series', async () => {
    const hidden = await roundTrip({
      kind: 'line',
      categories: ['A', 'B'],
      series: [{ name: 'Trend', values: [3, 1], markerSymbol: 'none' }],
    });
    expect(sliceElement(hidden.xml, 'ser')).toContain(
      '<c:marker><c:symbol val="none"/></c:marker>',
    );
    expect(hidden.spec.series[0]!.markerColor).toBeUndefined();
    expectSchemaValid(hidden.xml, 'chart');
    const bar = await roundTrip({
      kind: 'column',
      categories: ['A', 'B'],
      series: [{ name: 'Count', values: [3, 1] }],
    });
    expect(bar.xml).not.toContain('<c:marker>');
  });

  it('rejects an invalid markerColor before writing', async () => {
    await expect(
      roundTrip({
        kind: 'line',
        categories: ['A'],
        series: [{ name: 'a', values: [1], markerColor: 'red' }],
      }),
    ).rejects.toThrow(/series\[0\]\.markerColor.*invalid chart color/);
    await expect(
      roundTrip({
        kind: 'line',
        categories: ['A'],
        series: [{ name: 'a', values: [1], markerLineColor: 'red' }],
      }),
    ).rejects.toThrow(/series\[0\]\.markerLineColor.*invalid chart color/);
  });
});

describe('fn API: series lineColor', () => {
  skipIfNoXmllint('round-trips a slice border color on a doughnut series', async () => {
    const { spec, xml } = await roundTrip({
      kind: 'doughnut',
      categories: ['A', 'B'],
      series: [{ name: 'Share', values: [3, 1], lineWidthEmu: 9525, lineColor: '#F9F9F9' }],
    });
    expect(spec.series[0]).toMatchObject({ lineWidthEmu: 9525, lineColor: '#F9F9F9' });
    expect(spec.series[0]!.color).not.toBe('#F9F9F9');
    expect(sliceElement(xml, 'ser')).toContain(
      '<a:ln w="9525"><a:solidFill><a:srgbClr val="F9F9F9"/></a:solidFill></a:ln>',
    );
    expectSchemaValid(xml, 'chart');
  });

  it('writes the line in the series color when lineColor is unset', async () => {
    const { spec, xml } = await roundTrip({
      kind: 'line',
      categories: ['A', 'B'],
      series: [{ name: 'Trend', values: [3, 1], color: '#112233' }],
    });
    expect(sliceElement(xml, 'ser')).toContain(
      '<a:ln><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln>',
    );
    expect(spec.series[0]).toMatchObject({ color: '#112233', lineColor: '#112233' });
  });

  it('rejects an invalid lineColor before writing', async () => {
    await expect(
      roundTrip({
        kind: 'line',
        categories: ['A'],
        series: [{ name: 'a', values: [1], lineColor: 'red' }],
      }),
    ).rejects.toThrow(/series\[0\]\.lineColor.*invalid chart color/);
  });
});

describe('fn API: secondaryValueAxis color validation', () => {
  it.each([
    ['lineColor', { lineColor: 'not-a-color' }],
    ['majorGridlineColor', { majorGridlineColor: '#12345' }],
  ] as const)(
    'rejects an invalid secondary %s before writing',
    async (_field, secondaryValueAxis) => {
      const spec = (
        valueAxisLineColor: string | undefined,
        secondary: ChartSecondaryValueAxis | undefined,
      ): ChartSpec => ({
        kind: 'column',
        categories: ['A'],
        series: [
          { name: 'a', values: [1] },
          { name: 'b', values: [2], secondaryAxis: true },
        ],
        ...(valueAxisLineColor !== undefined ? { valueAxisLineColor } : {}),
        ...(secondary !== undefined ? { secondaryValueAxis: secondary } : {}),
      });
      const bad = Object.values(secondaryValueAxis)[0]!;
      await expect(roundTrip(spec(bad, undefined))).rejects.toThrow(/invalid chart color/);
      await expect(roundTrip(spec(undefined, secondaryValueAxis))).rejects.toThrow(
        /secondaryValueAxis\..*invalid chart color/,
      );
    },
  );
});

describe('fn API: chart axis positions', () => {
  it('puts the category axis on the left and the value axis at the bottom for horizontal bars', async () => {
    const { xml } = await roundTrip({
      kind: 'bar',
      categories: ['A', 'B'],
      series: [{ name: 'Count', values: [1, 2] }],
    });
    expect(sliceElement(xml, 'catAx')).toContain('<c:axPos val="l"/>');
    expect(sliceElement(xml, 'valAx')).toContain('<c:axPos val="b"/>');
  });

  skipIfNoXmllint('keeps the secondary axis on the right for horizontal bars', async () => {
    const { xml } = await roundTrip({
      kind: 'bar',
      categories: ['A', 'B'],
      series: [
        { name: 'Count', values: [1, 2] },
        { name: 'Share', values: [3, 4], secondaryAxis: true },
      ],
    });
    const axPositions = [...xml.matchAll(/<c:axPos val="(\w)"\/>/g)].map((m) => m[1]);
    // The deleted companion category axis follows the primary one (pptxgenjs too).
    expect(axPositions).toEqual(['l', 'b', 'r', 'l']);
    expectSchemaValid(xml, 'chart');
  });

  it('keeps the bottom / left layout for columns', async () => {
    const { xml } = await roundTrip({
      kind: 'column',
      categories: ['A', 'B'],
      series: [{ name: 'Count', values: [1, 2] }],
    });
    expect(sliceElement(xml, 'catAx')).toContain('<c:axPos val="b"/>');
    expect(sliceElement(xml, 'valAx')).toContain('<c:axPos val="l"/>');
  });
});

describe('fn API: chart point counts', () => {
  it('keeps the value array length when points are empty', async () => {
    const { spec } = await roundTrip({
      kind: 'column',
      categories: ['A', 'B', 'C'],
      series: [
        { name: 'Sparse', values: [1, null, 3] },
        { name: 'Empty', values: [null, null, null] },
        { name: 'Trailing', values: [1, null, null] },
      ],
    });
    expect(spec.series.map((s) => s.values)).toEqual([
      [1, null, 3],
      [null, null, null],
      [1, null, null],
    ]);
  });
});
