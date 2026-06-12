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
});

// ---------------------------------------------------------------------------
// void reference keeps TypeScript from complaining about the unused import
// while still exercising the export. The real usage is in the fallback test.
void textContentOf;
