// Real pptxgenjs 4.0.1 decks (test/fixtures/pptxgenjs) read into the public
// DTOs, written again through the public API into a fresh deck, and read
// back: the DTOs must match and the written XML must pass the schema. The raw
// XML checks cover what the reader cannot see (or would misread both ways).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';
import {
  type ChartSpec,
  type ParagraphSpec,
  type PresentationData,
  type ShapeParagraphElement,
  type SlideData,
  type SlideShapeData,
  addSlide,
  addSlideChart,
  addSlideTable,
  addSlideTextBox,
  createPresentation,
  findSlideLayoutByType,
  getParagraphAlignment,
  getShapeChartSpec,
  getShapeParagraphCount,
  getShapeParagraphElements,
  getSlideLayouts,
  getSlideShapes,
  getSlides,
  getTableCellParagraphs,
  getTableCells,
  inches,
  loadPresentation,
  readPackagePart,
  savePresentation,
  setShapeParagraphs,
  setTableCellParagraphs,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/pptxgenjs/${name}`, import.meta.url));

const decoder = new TextDecoder();
const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;
const xsd = (xml: string, schema: 'chart' | 'pml'): void => {
  if (isSchemaValidationAvailable()) expectSchemaValid(xml, schema);
};

const load = async (name: string): Promise<PresentationData> =>
  loadPresentation(await readFile(fixture(name)));

const freshSlide = (): { pres: PresentationData; slide: SlideData } => {
  const pres = createPresentation();
  const layout = findSlideLayoutByType(pres, 'blank') ?? getSlideLayouts(pres)[0]!;
  return { pres, slide: addSlide(pres, { layout }) };
};

const chartOf = (pres: PresentationData): ChartSpec => {
  const shape = getSlideShapes(getSlides(pres)[0]!).at(-1)!;
  return getShapeChartSpec(shape)!;
};

const partXml = (pres: PresentationData, name: string): string =>
  decoder.decode(readPackagePart(pres, name)!);

interface ChartRoundTrip {
  readonly before: ChartSpec;
  readonly after: ChartSpec;
  readonly oldXml: string;
  readonly newXml: string;
}

const sourceChartXml = (pres: PresentationData): string => {
  const slide = partXml(pres, '/ppt/slides/_rels/slide1.xml.rels');
  const target = /Target="[^"]*\/charts\/(chart\d+\.xml)"/.exec(slide)![1]!;
  return partXml(pres, `/ppt/charts/${target}`);
};

const valuesOf = (xml: string, tag: string): string[] =>
  [...xml.matchAll(new RegExp(`<c:${tag} val="([^"]*)"/>`, 'g'))].map((m) => m[1]!);

const axesOf = (xml: string): string[] =>
  [...xml.matchAll(/<c:(catAx|valAx)>.*?<\/c:\1>/gs)].map((m) => m[0]);

// Tick labels of a deleted axis are never drawn, so only live axes compare.
const axisLayout = (xml: string): Array<{ axPos: string; tickLblPos: string | undefined }> =>
  axesOf(xml).map((axis) => ({
    axPos: valuesOf(axis, 'axPos')[0]!,
    tickLblPos: axis.includes('<c:delete val="1"/>') ? undefined : valuesOf(axis, 'tickLblPos')[0],
  }));

// Per axis: the `w` of each gridline / axis-line <a:ln> outside title and labels.
// A deleted axis draws nothing, so its line is skipped.
const axisLineWidths = (xml: string): string[][] =>
  axesOf(xml)
    .filter((axis) => !axis.includes('<c:delete val="1"/>'))
    .map((axis) =>
      [
        ...axis.replace(/<c:(title|txPr)>.*?<\/c:\1>/gs, '').matchAll(/<a:ln(?: w="(\d+)")?[ >/]/g),
      ].map((m) => m[1] ?? 'none'),
    );

const axisTitleBodyPrs = (xml: string): string[] =>
  axesOf(xml).flatMap((axis) =>
    [...axis.matchAll(/<c:title>.*?(<a:bodyPr[^>]*>)/gs)].map((m) => m[1]!),
  );

const plotGroup = (xml: string, tag: string): string =>
  new RegExp(`<c:${tag}>.*?</c:${tag}>`, 's').exec(xml)![0];

const chartRoundTrip = async (name: string): Promise<ChartRoundTrip> => {
  const source = await load(name);
  const before = chartOf(source);
  const { pres, slide } = freshSlide();
  addSlideChart(slide, {
    x: inches(0.5),
    y: inches(0.5),
    w: inches(8),
    h: inches(4.5),
    spec: before,
  });
  const saved = await savePresentation(pres);
  const reloaded = await loadPresentation(saved);
  const newXml = partXml(reloaded, '/ppt/charts/chart1.xml');
  xsd(newXml, 'chart');
  return { before, after: chartOf(reloaded), oldXml: sourceChartXml(source), newXml };
};

describe('pptxgenjs compatibility: charts', () => {
  it('round-trips a column + line combo with a secondary axis', async () => {
    const { before, after, oldXml, newXml } = await chartRoundTrip('combo.pptx');
    expect(before.series[1]).toMatchObject({
      chartKind: 'line',
      secondaryAxis: true,
      markerSymbol: 'circle',
      markerSizePt: 6,
      smooth: false,
    });
    expect(before.valueAxisTitleRotationDeg).toBeUndefined();
    expect(after.valueAxisTitleRotationDeg).toBeUndefined();
    // pptxgenjs writes <c:showLeaderLines val="0"/> on every label group.
    expect(before.dataLabels?.showLeaderLines).toBe(false);
    expect(before.series.map((s) => s.dataLabels?.showLeaderLines)).toEqual([false, false]);
    // pptxgenjs picks the marker fill by the deck-wide series index and the
    // line + marker outline by the per-group one, so this fixture's marker is
    // filled in another color than its outline and line.
    expect(before.series[1]).toMatchObject({
      color: '#C0504D',
      markerColor: '#4F81BD',
      markerLineColor: '#C0504D',
    });
    expect(after).toEqual(before);

    // An axis title keeps PowerPoint's default orientation: no rot, no vert.
    const titles = axisTitleBodyPrs(newXml);
    expect(titles).toHaveLength(axisTitleBodyPrs(oldXml).length);
    expect(titles.length).toBeGreaterThan(0);
    for (const bodyPr of titles) expect(bodyPr).not.toMatch(/\b(rot|vert)=/);
    expect(valuesOf(plotGroup(newXml, 'barChart'), 'grouping')).toEqual(['clustered']);
    expect(valuesOf(plotGroup(newXml, 'lineChart'), 'grouping')).toEqual(['standard']);
    // The marker keeps its authored fill + outline instead of the theme default.
    expect(plotGroup(newXml, 'lineChart')).toContain(
      '<c:size val="6"/><c:spPr><a:solidFill><a:srgbClr val="4F81BD"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="C0504D"/>',
    );
    expect(valuesOf(newXml, 'showLeaderLines')).toEqual(valuesOf(oldXml, 'showLeaderLines'));
    expect(valuesOf(newXml, 'showLeaderLines').every((v) => v === '0')).toBe(true);
    expect(axisLayout(newXml)).toEqual(axisLayout(oldXml));
    // pptxgenjs draws axis lines and gridlines at 1 pt; PowerPoint's default
    // without `w` is 0.75 pt, so the widths must survive.
    expect(before).toMatchObject({
      valueAxisMajorGridlineWidthEmu: 12700,
      valueAxisLineWidthEmu: 12700,
      categoryAxisLineWidthEmu: 12700,
      secondaryValueAxis: { lineWidthEmu: 12700 },
    });
    expect(axisLineWidths(newXml)).toEqual(axisLineWidths(oldXml));
    expect(axisLineWidths(newXml).flat()).toContain('12700');
  });

  it('round-trips a horizontal bar chart with a secondary axis', async () => {
    const { before, after, oldXml, newXml } = await chartRoundTrip('bar-secondary.pptx');
    expect(before.kind).toBe('bar');
    expect(before.series[1]!.secondaryAxis).toBe(true);
    expect(after).toEqual(before);
    expect(axisLayout(newXml)).toEqual(axisLayout(oldXml));
    expect(axisLayout(newXml).map((axis) => axis.axPos)).toEqual(['l', 'b', 'r', 'l']);
    expect(axisLineWidths(newXml)).toEqual(axisLineWidths(oldXml));
    expect(valuesOf(newXml, 'showLeaderLines')).toEqual(valuesOf(oldXml, 'showLeaderLines'));
    for (const bodyPr of axisTitleBodyPrs(newXml)) expect(bodyPr).not.toMatch(/\b(rot|vert)=/);
  });

  it('round-trips a doughnut with series and per-point labels', async () => {
    const { before, after, oldXml, newXml } = await chartRoundTrip('doughnut.pptx');
    expect(before.series[0]!.pointDataLabels).toHaveLength(3);
    // The near-white slice border survives; pptxgenjs leaves the series fill
    // to the palette and the builder writes its default accent, which reads
    // back as an authored color.
    expect(before.series[0]!.lineColor).toBe('#F9F9F9');
    expect(before.series[0]!.dataLabels?.showLeaderLines).toBe(false);
    expect(after).toEqual({ ...before, series: [{ ...before.series[0], color: '#4472C4' }] });
    expect(plotGroup(newXml, 'doughnutChart')).toContain(
      '<a:ln w="9525"><a:solidFill><a:srgbClr val="F9F9F9"/></a:solidFill>',
    );
    expect(valuesOf(newXml, 'showLeaderLines')).toEqual(valuesOf(oldXml, 'showLeaderLines'));
  });
});

interface ParagraphDto {
  readonly align: ReturnType<typeof getParagraphAlignment>;
  readonly elements: ReadonlyArray<ShapeParagraphElement>;
}

const shapeParagraphs = (shape: SlideShapeData): ParagraphDto[] =>
  Array.from({ length: getShapeParagraphCount(shape) }, (_, i) => ({
    align: getParagraphAlignment(shape, i),
    elements: getShapeParagraphElements(shape, i),
  }));

const toSpecs = (paragraphs: ReadonlyArray<ParagraphDto>): ParagraphSpec[] =>
  paragraphs.map((p) => ({
    ...(p.align !== null ? { align: p.align } : {}),
    runs: p.elements.map((e) => {
      if (e.kind !== 'r') throw new Error(`fixture has a ${e.kind} element`);
      return { text: e.text, ...(e.format !== null ? { format: e.format } : {}) };
    }),
  }));

describe('pptxgenjs compatibility: text', () => {
  it('round-trips a text box and an autoPage table cell', async () => {
    const src = await load('text-table.pptx');
    const [textShape, tableShape] = getSlideShapes(getSlides(src)[0]!);
    const beforeText = shapeParagraphs(textShape!);
    const beforeCells = getTableCells(tableShape!).map((row) =>
      row.map((cell) => getTableCellParagraphs(cell)),
    );
    // pptxgenjs writes a raw CR LF inside <a:t>; the parser reads it as LF.
    expect(beforeText[0]!.elements.map((e) => (e.kind === 'r' ? e.text : e.kind))).toEqual([
      'Bold ',
      '\n',
      'plain',
    ]);

    const { pres, slide } = freshSlide();
    const text = addSlideTextBox(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(8),
      h: inches(1),
      text: '',
    });
    setShapeParagraphs(text, toSpecs(beforeText));
    const table = addSlideTable(slide, {
      x: inches(0.5),
      y: inches(2),
      w: inches(8),
      h: inches(3),
      rows: beforeCells.map((row) => row.map(() => '')),
    });
    getTableCells(table).forEach((row, r) => {
      row.forEach((cell, c) => setTableCellParagraphs(cell, toSpecs(beforeCells[r]![c]!)));
    });

    const reloaded = await loadPresentation(await savePresentation(pres));
    xsd(partXml(reloaded, '/ppt/slides/slide1.xml'), 'pml');
    const [againText, againTable] = getSlideShapes(getSlides(reloaded)[0]!);
    expect(shapeParagraphs(againText!)).toEqual(beforeText);
    expect(
      getTableCells(againTable!).map((row) => row.map((cell) => getTableCellParagraphs(cell))),
    ).toEqual(beforeCells);
  });
});

void skipIfNoXmllint;
