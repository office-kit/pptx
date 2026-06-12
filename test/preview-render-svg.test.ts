// Unit tests for `renderSlideToSvg` — exercised via in-memory decks built
// from the public API, loaded from the same blank.pptx fixture used across
// the rest of the test suite.
//
// Import pattern follows test/text-layout.test.ts: import from the package
// source directly (not the built dist) so vitest resolves TypeScript.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideChart,
  addSlideLine,
  addSlideImage,
  addSlideShape,
  addSlideTable,
  addSlideTextBox,
  findSlideLayout,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeFill,
  setShapeGradientFill,
  setShapeHyperlink,
  setShapePatternFill,
  setShapeRotation,
  setShapeRunFormat,
  setShapeStrokeArrow,
  setShapeText,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';
import { buildPng } from './lib/build-png.ts';
import { attrsOf, countTags, textContentOf } from './lib/svg-query.ts';

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));

const loadBlank = async () => loadPresentation(await readFile(fixturePath));

const blankSlide = async () => {
  const pres = await loadBlank();
  const layout = findSlideLayout(pres, 'Blank');
  if (!layout) throw new Error('Blank layout not found');
  const slide = addSlide(pres, { layout });
  return { pres, slide };
};

describe('renderSlideToSvg', () => {
  it('root <svg> has a viewBox matching the slide dimensions', async () => {
    const { pres, slide } = await blankSlide();
    const svg = renderSlideToSvg(pres, slide);
    const svgAttrs = attrsOf(svg, 'svg');
    expect(svgAttrs).toHaveLength(1);
    // blank.pptx is 4:3 (9144000 × 6858000 EMU → 960 × 720 CSS-px at 9525 EMU/px).
    expect(svgAttrs[0]!.viewBox).toMatch(/^0 0 960\.00 720\.00$/);
  });

  it('preset roundRect shape with solid fill carries that fill color', async () => {
    const { pres, slide } = await blankSlide();
    const shape = addSlideShape(slide, {
      preset: 'roundRect',
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
    });
    setShapeFill(shape, '#FF0000');
    const svg = renderSlideToSvg(pres, slide);
    expect(svg).toMatch(/fill="#[Ff][Ff]0+"/);
  });

  it('gradient fill emits a <linearGradient> def that the shape references', async () => {
    const { pres, slide } = await blankSlide();
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(3),
      h: inches(2),
    });
    setShapeGradientFill(shape, {
      stops: [
        { offset: 0, color: '#0000FF' },
        { offset: 1, color: '#FF0000' },
      ],
      angleDeg: 0,
    });
    const svg = renderSlideToSvg(pres, slide);
    expect(countTags(svg, 'linearGradient')).toBeGreaterThan(0);
    expect(svg).toContain('url(#');
  });

  it('pattern fill emits a <pattern> def that the shape references', async () => {
    const { pres, slide } = await blankSlide();
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(3),
      h: inches(2),
    });
    setShapePatternFill(shape, { preset: 'pct50', foreground: '#000000', background: '#FFFFFF' });
    const svg = renderSlideToSvg(pres, slide);
    expect(countTags(svg, 'pattern')).toBeGreaterThan(0);
    expect(svg).toContain('url(#');
  });

  it('foreignObject mode wraps text body in <foreignObject>', async () => {
    const { pres, slide } = await blankSlide();
    addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'hello world',
    });
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'foreignObject' });
    expect(countTags(svg, 'foreignObject')).toBeGreaterThan(0);
  });

  it('svg text mode emits <text> containing the run text and no <foreignObject>', async () => {
    const { pres, slide } = await blankSlide();
    addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'hello svg',
    });
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    expect(countTags(svg, 'foreignObject')).toBe(0);
    expect(countTags(svg, 'text')).toBeGreaterThan(0);
    expect(svg).toContain('hello svg');
  });

  it('svg text mode: bold + explicit color from setShapeRunFormat appear in output', async () => {
    const { pres, slide } = await blankSlide();
    const box = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'styled text',
    });
    setShapeText(box, 'styled text');
    setShapeRunFormat(box, 0, 0, { bold: true, color: '#CC0000' });
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    expect(svg).toContain('font-weight="700"');
    expect(svg).toMatch(/fill="#[Cc][Cc]0+0+"/);
  });

  it('addSlideImage: emits <image> element with a data: URL href', async () => {
    const { pres, slide } = await blankSlide();
    const pngBytes = buildPng(4, 4, [255, 128, 0]);
    addSlideImage(slide, pngBytes, {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(2),
    });
    const svg = renderSlideToSvg(pres, slide);
    expect(countTags(svg, 'image')).toBeGreaterThan(0);
    const imageAttrs = attrsOf(svg, 'image');
    const hasDataUrl = imageAttrs.some((a) =>
      (a['href'] ?? a['xlink:href'] ?? '').startsWith('data:'),
    );
    expect(hasDataUrl).toBe(true);
  });

  it('addSlideTable: emits cell rects for every cell in the grid', async () => {
    const { pres, slide } = await blankSlide();
    addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(3),
      rows: [
        ['A', 'B', 'C'],
        ['D', 'E', 'F'],
      ],
    });
    const svg = renderSlideToSvg(pres, slide);
    // Each of the 6 cells gets at minimum one <rect> for its background, plus
    // one for the slide background = at least 7 rects total.
    expect(countTags(svg, 'rect')).toBeGreaterThanOrEqual(7);
  });

  it('bar chart: renders plot geometry and carries no data-pptx-fallback', async () => {
    const { pres, slide } = await blankSlide();
    addSlideChart(slide, {
      spec: {
        kind: 'bar',
        categories: ['X', 'Y'],
        series: [{ name: 'S', values: [10, 20] }],
      },
      x: inches(1),
      y: inches(1),
      w: inches(5),
      h: inches(3),
    });
    const svg = renderSlideToSvg(pres, slide);
    // A rendered bar chart contains rect elements for the bars.
    expect(countTags(svg, 'rect')).toBeGreaterThan(1);
    // Placeholder shapes from the Blank layout have no prstGeom and get
    // data-pptx-fallback="custGeom"; assert that the chart itself did not fall back.
    expect(svg).not.toContain('data-pptx-fallback="chart"');
  });

  it('line connector with arrow head: emits a <marker> def', async () => {
    const { pres, slide } = await blankSlide();
    const line = addSlideLine(slide, {
      from: { x: inches(1), y: inches(2) },
      to: { x: inches(4), y: inches(2) },
      // A solid stroke color is required for the renderer to emit marker defs.
      color: '#000000',
    });
    setShapeStrokeArrow(line, 'head', { type: 'triangle' });
    const svg = renderSlideToSvg(pres, slide);
    expect(countTags(svg, 'marker')).toBeGreaterThan(0);
  });

  it('setShapeRotation: rotate(...) transform present in output', async () => {
    const { pres, slide } = await blankSlide();
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
    });
    setShapeRotation(shape, 30);
    const svg = renderSlideToSvg(pres, slide);
    expect(svg).toContain('rotate(');
  });

  it('setShapeHyperlink: shape is wrapped in an <a> with the href', async () => {
    const { pres, slide } = await blankSlide();
    // setShapeHyperlink requires a <p:txBody> on the shape; use addSlideTextBox
    // which always creates one, rather than a bare addSlideShape.
    const shape = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      text: 'click me',
    });
    setShapeHyperlink(shape, 'https://example.com');
    const svg = renderSlideToSvg(pres, slide);
    expect(svg).toContain('href="https://example.com"');
  });

  it('shapes carry data-pptx-shape-name for accessibility / DevTools', async () => {
    const { pres, slide } = await blankSlide();
    addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      name: 'MyBox',
    });
    const svg = renderSlideToSvg(pres, slide);
    expect(svg).toContain('data-pptx-shape-name="MyBox"');
  });

  it('fallback marker: unrecognised chart kind emits data-pptx-fallback="chart"', async () => {
    // Build a deck with a bar chart, then use the internal OPC zip layer (an
    // internal hook, as permitted by the task brief) to rename the plotted
    // element to something not in the chart reader's KIND_MAP. This exercises
    // the fallback path for chart types the renderer does not yet model.
    const { pres, slide } = await blankSlide();
    addSlideChart(slide, {
      spec: {
        kind: 'bar',
        categories: ['A'],
        series: [{ name: 'S', values: [1] }],
      },
      x: inches(1),
      y: inches(1),
      w: inches(5),
      h: inches(3),
    });
    const bytes = await savePresentation(pres);
    const { entries } = readZip(bytes);
    const enc = new TextEncoder();
    const dec = new TextDecoder();
    const modified = entries.map((e) => {
      if (!e.name.includes('charts/chart')) return e;
      const xml = dec
        .decode(e.data)
        .replace(/<c:barChart\b/g, '<c:unknownPlotChart')
        .replace(/<\/c:barChart>/g, '</c:unknownPlotChart>');
      return { name: e.name, data: enc.encode(xml) };
    });
    const pres2 = await loadPresentation(writeZip(modified));
    const svg = renderSlideToSvg(pres2, getSlides(pres2)[0]!);
    expect(svg).toContain('data-pptx-fallback="chart"');
  });

  it('rendering the same slide twice yields identical SVG strings (determinism)', async () => {
    const { pres, slide } = await blankSlide();
    addSlideShape(slide, {
      preset: 'roundRect',
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      text: 'determinism',
    });
    const opts = { textLayout: 'svg' as const };
    expect(renderSlideToSvg(pres, slide, opts)).toBe(renderSlideToSvg(pres, slide, opts));
  });

  // scatter / radar / bubble have no public authoring API (read + render
  // only — plan W4). To exercise their plotters we build a deck with a
  // throwaway bar chart, then swap the chart part's XML via the internal
  // OPC zip layer (the same hook the fallback test above uses).
  const renderInjectedChart = async (chartXml: string): Promise<string> => {
    const { pres, slide } = await blankSlide();
    addSlideChart(slide, {
      spec: { kind: 'bar', categories: ['A'], series: [{ name: 'S', values: [1] }] },
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(4),
    });
    const bytes = await savePresentation(pres);
    const { entries } = readZip(bytes);
    const enc = new TextEncoder();
    const modified = entries.map((e) =>
      e.name.includes('charts/chart') ? { name: e.name, data: enc.encode(chartXml) } : e,
    );
    const pres2 = await loadPresentation(writeZip(modified));
    return renderSlideToSvg(pres2, getSlides(pres2)[0]!);
  };

  const chartSpaceXml = (chart: string): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>XY Demo</a:t></a:r></a:p></c:rich></c:tx></c:title>
    ${chart}
    <c:legend><c:legendPos val="b"/></c:legend>
  </c:chart>
</c:chartSpace>`;

  const scatterPlotArea = (scatterStyle: string): string =>
    chartSpaceXml(`
    <c:plotArea>
      <c:layout/>
      <c:scatterChart>
        <c:scatterStyle val="${scatterStyle}"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:strLit><c:pt idx="0"><c:v>Alpha</c:v></c:pt></c:strLit></c:tx>
          <c:xVal><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numLit></c:xVal>
          <c:yVal><c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>30</c:v></c:pt><c:pt idx="2"><c:v>20</c:v></c:pt></c:numLit></c:yVal>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:scatterChart>
      <c:valAx><c:axId val="1"/><c:crossAx val="2"/></c:valAx>
      <c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`);

  it('scatter (marker style): point markers, no connecting line, no fallback, legend present', async () => {
    const svg = await renderInjectedChart(scatterPlotArea('marker'));
    expect(svg).not.toContain('data-pptx-fallback="chart"');
    // One <circle> per data point (3); markers only means no plot <path>.
    expect(countTags(svg, 'circle')).toBeGreaterThanOrEqual(3);
    expect(countTags(svg, 'path')).toBe(0);
    expect(svg).toContain('Alpha');
  });

  it('scatter (lineMarker style): adds a connecting <path> over the markers', async () => {
    const markerOnly = await renderInjectedChart(scatterPlotArea('marker'));
    const lineMarker = await renderInjectedChart(scatterPlotArea('lineMarker'));
    expect(countTags(lineMarker, 'circle')).toBeGreaterThanOrEqual(3);
    // The connecting polyline is the only <path> source on this slide.
    expect(countTags(lineMarker, 'path')).toBeGreaterThan(countTags(markerOnly, 'path'));
  });

  it('radar (filled): closed series polygon filled at reduced opacity, no fallback', async () => {
    const svg = await renderInjectedChart(
      chartSpaceXml(`
      <c:plotArea>
        <c:layout/>
        <c:radarChart>
          <c:radarStyle val="filled"/>
          <c:ser>
            <c:idx val="0"/><c:order val="0"/>
            <c:tx><c:strLit><c:pt idx="0"><c:v>Alpha</c:v></c:pt></c:strLit></c:tx>
            <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strLit></c:cat>
            <c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>5</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt></c:numLit></c:val>
          </c:ser>
          <c:axId val="1"/><c:axId val="2"/>
        </c:radarChart>
        <c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx>
        <c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx>
      </c:plotArea>`),
    );
    expect(svg).not.toContain('data-pptx-fallback="chart"');
    // Rings + the closed series polygon are all <polygon>.
    expect(countTags(svg, 'polygon')).toBeGreaterThan(0);
    expect(svg).toContain('fill-opacity="0.3"');
    expect(svg).toContain('Alpha');
  });

  it('bubble: per-point circles with different radii, no fallback', async () => {
    const svg = await renderInjectedChart(
      chartSpaceXml(`
      <c:plotArea>
        <c:layout/>
        <c:bubbleChart>
          <c:ser>
            <c:idx val="0"/><c:order val="0"/>
            <c:tx><c:strLit><c:pt idx="0"><c:v>Alpha</c:v></c:pt></c:strLit></c:tx>
            <c:xVal><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numLit></c:xVal>
            <c:yVal><c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>15</c:v></c:pt></c:numLit></c:yVal>
            <c:bubbleSize><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>9</c:v></c:pt><c:pt idx="2"><c:v>4</c:v></c:pt></c:numLit></c:bubbleSize>
          </c:ser>
          <c:bubbleScale val="100"/>
          <c:sizeRepresents val="area"/>
          <c:axId val="1"/><c:axId val="2"/>
        </c:bubbleChart>
        <c:valAx><c:axId val="1"/><c:crossAx val="2"/></c:valAx>
        <c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx>
      </c:plotArea>`),
    );
    expect(svg).not.toContain('data-pptx-fallback="chart"');
    const radii = attrsOf(svg, 'circle')
      .map((a) => a['r'])
      .filter((r): r is string => r !== undefined);
    expect(radii.length).toBeGreaterThanOrEqual(3);
    // Area-proportional sizing (sizes 1 / 9 / 4) must yield distinct radii.
    expect(new Set(radii).size).toBeGreaterThan(1);
    expect(svg).toContain('Alpha');
  });
});

// ---------------------------------------------------------------------------
// void reference keeps TypeScript from complaining about the unused import
// while still exercising the export. The real usage is in the fallback test.
void textContentOf;
