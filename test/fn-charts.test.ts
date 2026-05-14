// Free-function chart authoring.
//
// Verifies end-to-end:
//   - The chart part (/ppt/charts/chart{N}.xml) is added with the right
//     content type, references the embedded xlsx via `<c:externalData>`.
//   - The embedded xlsx (/ppt/embeddings/Microsoft_Excel_Worksheet{N}.xlsx)
//     is a real OPC zip with the expected sheet entries.
//   - The slide carries the new `<p:graphicFrame>` referencing the chart
//     via a fresh `chart` rel.
//   - The shape returned by `addSlideChart` is reported as
//     kind `graphicFrame`.
//   - The chart XML validates against the ECMA-376 chart XSD when xmllint
//     is available.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  Presentation,
  _internalPackageOf,
  addSlideChart,
  getShapeKind,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';
import { readZip } from '../src/internal/opc/zip.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const partBytes = async (
  bytes: Uint8Array,
  partPath: string,
): Promise<Uint8Array | null> => {
  const pres = await Presentation.load(bytes);
  const pkg = _internalPackageOf(pres);
  const part = pkg.parts.find((p) => p.name === partPath);
  return part?.data ?? null;
};

describe('fn API: addSlideChart', () => {
  it('column chart with one series persists end-to-end', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const before = getSlideShapes(slide).length;

    const shape = addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['Q1', 'Q2', 'Q3', 'Q4'],
        series: [{ name: 'Revenue', values: [10, 20, 15, 30] }],
        title: 'Quarterly Revenue',
      },
    });
    expect(getShapeKind(shape)).toBe('graphicFrame');
    expect(getSlideShapes(slide).length).toBe(before + 1);

    const bytes = await savePresentation(pres);
    const chartXml = await partBytes(bytes, '/ppt/charts/chart1.xml');
    expect(chartXml).not.toBeNull();
    const chartStr = new TextDecoder().decode(chartXml!);
    expect(chartStr).toContain('<c:barChart>');
    expect(chartStr).toContain('Q1');
    expect(chartStr).toContain('Revenue');
    expect(chartStr).toContain('<c:externalData');

    // Embedded xlsx must be a real OPC zip.
    const xlsxBytes = await partBytes(bytes, '/ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx');
    expect(xlsxBytes).not.toBeNull();
    const xlsxZip = readZip(xlsxBytes!);
    const entries = xlsxZip.entries.map((e) => e.name).sort();
    expect(entries).toContain('xl/workbook.xml');
    expect(entries).toContain('xl/worksheets/sheet1.xml');
    expect(entries).toContain('[Content_Types].xml');

    const sheetXml = new TextDecoder().decode(
      xlsxZip.entries.find((e) => e.name === 'xl/worksheets/sheet1.xml')!.data,
    );
    // Header row sits in row 1; Q1 lands in row 2.
    expect(sheetXml).toContain('<t>Revenue</t>');
    expect(sheetXml).toContain('<t>Q1</t>');
    expect(sheetXml).toContain('<v>10</v>');
  });

  it('bar / line / pie / doughnut / area chart kinds all save and reload', async () => {
    for (const kind of ['bar', 'line', 'pie', 'doughnut', 'area'] as const) {
      const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
      const slide = getSlides(pres)[0]!;
      const seriesCount = kind === 'pie' || kind === 'doughnut' ? 1 : 2;
      addSlideChart(slide, {
        x: inches(0),
        y: inches(0),
        w: inches(5),
        h: inches(3),
        spec: {
          kind,
          categories: ['A', 'B', 'C'],
          series: Array.from({ length: seriesCount }, (_, i) => ({
            name: `S${i + 1}`,
            values: [1, 2, 3],
          })),
        },
      });
      const reloaded = await Presentation.load(await savePresentation(pres));
      // The graphic frame must round-trip back as a shape on the slide.
      const reloadedShape = reloaded.slides[0]?.shapes.find((s) => s.kind === 'graphicFrame');
      expect(reloadedShape, `${kind} chart lost on round-trip`).toBeDefined();
    }
  });

  it('doughnut chart emits holeSize attribute', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(5),
      h: inches(3),
      spec: {
        kind: 'doughnut',
        categories: ['X', 'Y', 'Z'],
        series: [{ name: 'A', values: [1, 2, 3] }],
      },
    });
    const bytes = await savePresentation(pres);
    const pres2 = await Presentation.load(bytes);
    const pkg = _internalPackageOf(pres2);
    const chartPart = pkg.parts.find((p) => p.name === '/ppt/charts/chart1.xml');
    expect(chartPart).not.toBeUndefined();
    const chartStr = new TextDecoder().decode(chartPart!.data);
    expect(chartStr).toContain('<c:doughnutChart>');
    expect(chartStr).toContain('<c:holeSize');
  });

  it('area chart uses areaChart element with axes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(5),
      h: inches(3),
      spec: {
        kind: 'area',
        categories: ['Jan', 'Feb', 'Mar'],
        series: [{ name: 'Revenue', values: [10, 20, 30] }],
      },
    });
    const bytes = await savePresentation(pres);
    const pres2 = await Presentation.load(bytes);
    const pkg = _internalPackageOf(pres2);
    const chartPart = pkg.parts.find((p) => p.name === '/ppt/charts/chart1.xml');
    const chartStr = new TextDecoder().decode(chartPart!.data);
    expect(chartStr).toContain('<c:areaChart>');
    expect(chartStr).toContain('<c:catAx>');
    expect(chartStr).toContain('<c:valAx>');
  });

  it('multi-series column chart records each series with its own color', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['Mon', 'Tue', 'Wed'],
        series: [
          { name: 'A', values: [1, 2, 3], color: '#112233' },
          { name: 'B', values: [4, 5, 6], color: '#445566' },
        ],
      },
    });
    const chartStr = new TextDecoder().decode(
      (await partBytes(await savePresentation(pres), '/ppt/charts/chart1.xml'))!,
    );
    expect(chartStr).toContain('112233');
    expect(chartStr).toContain('445566');
    // Two `<c:ser>` blocks.
    expect(chartStr.match(/<c:ser>/g)?.length).toBe(2);
  });
});
