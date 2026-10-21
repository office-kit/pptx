// setChartData — replace a chart's values, keeping its formatting when the
// series stay the same.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import {
  type ChartSpec,
  type PresentationData,
  addSlideChart,
  getSlideCharts,
  getSlides,
  inches,
  loadPresentation,
  readPackagePart,
  savePresentation,
  setChartData,
} from '../src/api/index.ts';
import { INTERNAL_PACKAGE } from '../src/api/_internal-symbols.ts';
import { partName } from '../src/internal/opc/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const CHART_PART = '/ppt/charts/chart1.xml';
const WORKBOOK_PART = '/ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx';

const COLUMN: ChartSpec = {
  kind: 'column',
  title: 'Sales',
  valueAxisTitle: 'Units',
  categories: ['Q1', 'Q2'],
  series: [
    { name: 'East', values: [1, 2], color: '#C00000' },
    { name: 'West', values: [3, 4] },
  ],
};

async function deckWith(spec: ChartSpec): Promise<PresentationData> {
  const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
  addSlideChart(getSlides(pres)[0]!, {
    x: inches(0.5),
    y: inches(0.5),
    w: inches(6),
    h: inches(4),
    spec,
  });
  return loadPresentation(await savePresentation(pres));
}

const chartOf = (pres: PresentationData) => getSlideCharts(getSlides(pres)[0]!)[0]!;
const chartXml = (pres: PresentationData) =>
  new TextDecoder().decode(readPackagePart(pres, CHART_PART)!);

// Stands in for formatting another app wrote, which ChartSpec cannot carry.
function editChartXml(pres: PresentationData, edit: (xml: string) => string): void {
  const part = pres[INTERNAL_PACKAGE].getPart(partName(CHART_PART))!;
  part.data = new TextEncoder().encode(edit(new TextDecoder().decode(part.data)));
}

const MARKER = '<c:dPt><c:idx val="0"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/></c:dPt>';

function withForeignFormatting(pres: PresentationData): void {
  editChartXml(pres, (xml) =>
    xml
      .replace(/(<c:ser>(?:(?!<\/c:ser>)[\s\S])*?<\/c:tx>)/, `$1${MARKER}`)
      .replaceAll(
        /<c:numCache><c:formatCode>General<\/c:formatCode>/g,
        '<c:numCache><c:formatCode>#,##0"件"</c:formatCode>',
      ),
  );
}

function sheetXml(pres: PresentationData): string {
  const files = unzipSync(readPackagePart(pres, WORKBOOK_PART)!);
  return strFromU8(files['xl/worksheets/sheet1.xml']!);
}

describe('setChartData', () => {
  it('replaces values in place, keeping everything else in the chart', async () => {
    const pres = await deckWith(COLUMN);
    withForeignFormatting(pres);
    const before = chartXml(pres);

    setChartData(chartOf(pres), {
      categories: ['Q1', 'Q2', 'Q3'],
      series: [
        { name: 'West', values: [30, 40, null] },
        { name: 'East', values: [10, 20, 25] },
      ],
    });

    const reread = await loadPresentation(await savePresentation(pres));
    const spec = chartOf(reread).spec!;
    expect(spec.categories).toEqual(['Q1', 'Q2', 'Q3']);
    expect(spec.series.map((s) => [s.name, s.values])).toEqual([
      ['East', [10, 20, 25]],
      ['West', [30, 40, null]],
    ]);
    const after = chartXml(reread);
    expect(after).toContain(MARKER);
    expect(after.match(/#,##0"件"/g)).toHaveLength(2);
    // Only the caches and their ranges differ.
    const strip = (xml: string) =>
      xml.replace(/<c:(str|num)Cache>[\s\S]*?<\/c:\1Cache>/g, '').replace(/<c:f>[^<]*<\/c:f>/g, '');
    expect(strip(after)).toBe(strip(before));
    expect(after).toContain('<c:f>Sheet1!$A$2:$A$4</c:f>');
    expect(after).toContain('<c:f>Sheet1!$C$2:$C$4</c:f>');

    // The workbook "Edit data" opens holds the new values.
    const sheet = sheetXml(reread);
    expect(sheet).toContain('Q3');
    expect(sheet).toContain('<v>25</v>');
  });

  it('changes only values when categories and series are left out', async () => {
    const pres = await deckWith(COLUMN);
    setChartData(chartOf(pres), {
      series: [
        { name: 'East', values: [5, 6] },
        { name: 'West', values: [7, 8] },
      ],
    });
    const spec = chartOf(pres).spec!;
    expect(spec.categories).toEqual(['Q1', 'Q2']);
    expect(spec.series.map((s) => s.values)).toEqual([
      [5, 6],
      [7, 8],
    ]);
  });

  it('builds the chart again when the kind or the series change', async () => {
    const pres = await deckWith(COLUMN);
    withForeignFormatting(pres);
    setChartData(chartOf(pres), {
      kind: 'line',
      series: [
        { name: 'East', values: [1, 1] },
        { name: 'North', values: [2, 2] },
      ],
    });
    const spec = chartOf(pres).spec!;
    expect(spec.kind).toBe('line');
    expect(spec.title).toBe('Sales');
    expect(spec.valueAxisTitle).toBe('Units');
    expect(spec.series.map((s) => s.name)).toEqual(['East', 'North']);
    expect(spec.series[0]!.color).toBe('#C00000');
    expect(chartXml(pres)).not.toContain('件');

    // A series renamed is a series removed and another added.
    setChartData(chartOf(pres), {
      series: [
        { name: 'East', values: [1, 1] },
        { name: 'South', values: [2, 2] },
      ],
    });
    expect(chartOf(pres).spec!.series.map((s) => s.name)).toEqual(['East', 'South']);

    setChartData(chartOf(pres), { kind: 'pie', series: [{ name: 'Share', values: [3, 7] }] });
    expect(chartOf(pres).spec!.kind).toBe('pie');
    expect(() =>
      setChartData(chartOf(pres), {
        kind: 'doughnut',
        series: [
          { name: 'A', values: [1, 2] },
          { name: 'B', values: [1, 2] },
        ],
      }),
    ).toThrow(/a doughnut chart takes exactly one series/);
  });

  it('throws when a series does not have one value per category, changing nothing', async () => {
    const pres = await deckWith(COLUMN);
    const before = chartXml(pres);
    expect(() => setChartData(chartOf(pres), { categories: ['Q1', 'Q2', 'Q3'] })).toThrow(
      /series "East" has 2 values for 3 categories/,
    );
    expect(() =>
      setChartData(chartOf(pres), {
        series: [
          { name: 'East', values: [1] },
          { name: 'West', values: [1, 2] },
        ],
      }),
    ).toThrow(/series "East" has 1 values for 2 categories/);
    expect(() =>
      setChartData(chartOf(pres), {
        series: [
          { name: 'East', values: [1, 2] },
          { name: 'East', values: [1, 2] },
        ],
      }),
    ).toThrow(/distinct/);
    expect(chartXml(pres)).toBe(before);
  });

  it('throws on a chart whose values it cannot replace', async () => {
    const pres = await deckWith({
      kind: 'scatter',
      categories: [],
      series: [{ name: 'S', values: [1, 2], xValues: [3, 4] }],
    });
    expect(() => setChartData(chartOf(pres), { series: [] })).toThrow(
      /the data of a scatter chart cannot be replaced/,
    );
    editChartXml(pres, (xml) => xml.replace(/<c:scatterChart>[\s\S]*<\/c:scatterChart>/, ''));
    expect(() => setChartData(chartOf(pres), {})).toThrow(
      /the data of this chart cannot be replaced/,
    );
  });

  it('keeps a multi-level category axis but refuses to replace its categories', async () => {
    const pres = await deckWith({
      kind: 'column',
      categories: ['Q1', 'Q2'],
      categoryGroupLevels: [['2026', '']],
      series: [{ name: 'S', values: [1, 2] }],
    });
    setChartData(chartOf(pres), { series: [{ name: 'S', values: [3, 4] }] });
    const spec = chartOf(pres).spec!;
    expect(spec.series[0]!.values).toEqual([3, 4]);
    expect(spec.categoryGroupLevels).toEqual([['2026', '']]);
    expect(() => setChartData(chartOf(pres), { categories: ['A', 'B'] })).toThrow(
      /several category levels/,
    );
  });
});
