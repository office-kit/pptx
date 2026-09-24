// Unit tests for table-cell text fidelity in `renderSlideToSvg`.
//
// Cell text is authored through the public API (`setTableCellTextFormat` /
// `setTableCellAlignment`) and rendered in both text-layout modes, asserting
// that per-cell run format, alignment, the unstyled default size, wrapping,
// and the foreignObject path all reach the output.
//
// Import pattern follows test/preview-render-svg.test.ts: import from package
// source directly so vitest resolves TypeScript.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTable,
  findSlideLayout,
  getTableCell,
  inches,
  loadPresentation,
  setTableCellFill,
  clearTableCellFill,
  setTableStyleFlags,
  setTableCellAlignment,
  setTableCellAnchor,
  setTableCellTextDirection,
  setTableCellTextFormat,
} from '../src/api/index.ts';
import { renderSlideToSvg, measureTableCellHeight } from '../packages/preview/src/index.ts';
import { attrsOf, countTags, textContentOf } from './lib/svg-query.ts';

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));

const blankSlide = async () => {
  const pres = await loadPresentation(await readFile(fixturePath));
  const layout = findSlideLayout(pres, 'Blank');
  if (!layout) throw new Error('Blank layout not found');
  const slide = addSlide(pres, { layout });
  return { pres, slide };
};

describe('table cell text rendering', () => {
  it('renders authored baseline offsets in both text paths', async () => {
    const { pres, slide } = await blankSlide();
    const table = addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(1),
      rows: [['Up', 'Down']],
    });
    setTableCellTextFormat(getTableCell(table, 0, 0), { size: 30, baseline: 0.4 });
    setTableCellTextFormat(getTableCell(table, 0, 1), { size: 30, baseline: -0.2 });
    const html = renderSlideToSvg(pres, slide);
    expect(html).toContain('vertical-align:16.000px');
    expect(html).toContain('vertical-align:-8.000px');
    expect(html).toContain('font-size:26.000px');
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    expect(svg).toMatch(/baseline-shift="16(?:\.0+)?"/);
    expect(svg).toMatch(/baseline-shift="-8(?:\.0+)?"/);
    expect(svg).toMatch(/font-size="26(?:\.0+)?"/);
  });

  it('honors the kerning threshold in HTML and SVG cell text', async () => {
    const { pres, slide } = await blankSlide();
    const table = addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(1),
      rows: [['AV', 'AV', 'AV']],
    });
    setTableCellTextFormat(getTableCell(table, 0, 0), { size: 20, kern: 2000 });
    setTableCellTextFormat(getTableCell(table, 0, 1), { size: 19, kern: 2000 });
    setTableCellTextFormat(getTableCell(table, 0, 2), { size: 20, kern: 0 });
    for (const textLayout of ['foreignObject', 'svg'] as const) {
      const svg = renderSlideToSvg(pres, slide, { textLayout });
      expect((svg.match(/font-kerning:normal/g) ?? []).length).toBe(1);
      expect((svg.match(/font-kerning:none/g) ?? []).length).toBe(2);
    }
  });

  it('keeps exact EMU minimum heights free of floating-point ceiling growth', async () => {
    const { pres, slide } = await blankSlide();
    const table = addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(0.4),
      rows: [['']],
    });
    expect(measureTableCellHeight(pres, table, 0, 0)).toBe(365760);
  });

  it('explicit no fill overrides table style and clearing restores inherited shading', async () => {
    const { pres, slide } = await blankSlide();
    const table = addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      rows: [['A', 'B']],
    });
    setTableStyleFlags(table, { firstRow: true });
    const cell = getTableCell(table, 0, 0);
    const fill = () =>
      renderSlideToSvg(pres, slide).match(/<rect data-pptx-cell="0,0"[^>]* fill="([^"]+)"/)?.[1];
    const inherited = fill();
    expect(inherited).not.toBe('none');
    setTableCellFill(cell, 'FF0000');
    expect(fill()).toBe('#FF0000');
    setTableCellFill(cell, null);
    expect(fill()).toBe('none');
    clearTableCellFill(cell);
    expect(fill()).toBe(inherited);
  });

  it.each([
    ['vert', 'vertical-rl', 'sideways', 'rotate(90'],
    ['vert270', 'vertical-lr', 'sideways', 'rotate(270'],
    ['wordArtVert', 'vertical-rl', 'upright', null],
  ] as const)(
    'renders %s cell text in both layout modes',
    async (direction, mode, orientation, rotation) => {
      const { pres, slide } = await blankSlide();
      const table = addSlideTable(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(3),
        rows: [['AB', 'Neighbor']],
      });
      const cell = getTableCell(table, 0, 0);
      setTableCellTextDirection(cell, direction);
      setTableCellAnchor(cell, 'bottom');
      const html = renderSlideToSvg(pres, slide, { textLayout: 'foreignObject' });
      expect(html).toContain(`writing-mode:${mode};text-orientation:${orientation}`);
      expect(html).toContain('justify-content:flex-end');
      if (direction === 'vert270') expect(html).toContain('transform:rotate(180deg)');
      const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
      expect(textContentOf(svg)).toContain('Neighbor');
      expect(textContentOf(svg)).toContain('A');
      expect(textContentOf(svg)).toContain('B');
      if (rotation) expect(svg).toContain(rotation);
      else expect(svg).not.toMatch(/rotate\((?:90|270)/);
    },
  );
  it('svg mode: an explicitly formatted cell carries its size / weight / color per run', async () => {
    const { pres, slide } = await blankSlide();
    const table = addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(2),
      rows: [['Styled', 'Plain']],
    });
    // 28 pt bold red on the first cell; the sibling stays unformatted.
    setTableCellTextFormat(getTableCell(table, 0, 0), { size: 28, bold: true, color: '#CC0000' });

    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    // 28 pt → 28 * 96/72 = 37.33 px.
    expect(svg).toContain('font-size="37.33"');
    expect(svg).toContain('font-weight="700"');
    expect(svg).toMatch(/fill="#[Cc][Cc]0+0+"/);
    // Both cell texts reach the output.
    const text = textContentOf(svg);
    expect(text).toContain('Styled');
    expect(text).toContain('Plain');
  });

  it('svg mode: an unformatted cell falls back to the 18 pt table-cell default', async () => {
    const { pres, slide } = await blankSlide();
    addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(2),
      rows: [['Plain']],
    });
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    // 18 pt → 18 * 96/72 = 24 px, PowerPoint's default for a freshly
    // inserted table cell (no explicit <a:rPr sz>).
    expect(svg).toContain('font-size="24"');
  });

  it('svg mode: a centered cell renders its text with text-anchor="middle"', async () => {
    const { pres, slide } = await blankSlide();
    const table = addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(2),
      rows: [['Centered']],
    });
    setTableCellAlignment(getTableCell(table, 0, 0), 'center');
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    const anchors = attrsOf(svg, 'text').map((a) => a['text-anchor']);
    expect(anchors).toContain('middle');
  });

  it('svg mode: long cell text wraps within the cell width across multiple lines', async () => {
    const { pres, slide } = await blankSlide();
    // A single narrow cell forces the long sentence to wrap.
    addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(1.5),
      h: inches(2),
      rows: [['Wrapping cell text spans several lines here']],
    });
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    // One <text> element is emitted per laid-out line; wrapping yields >= 2.
    expect(countTags(svg, 'text')).toBeGreaterThanOrEqual(2);
  });

  it('foreignObject mode: cell text is emitted inside a <foreignObject>', async () => {
    const { pres, slide } = await blankSlide();
    const table = addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(2),
      rows: [['Styled', 'Plain']],
    });
    setTableCellTextFormat(getTableCell(table, 0, 0), { size: 28, bold: true, color: '#CC0000' });
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'foreignObject' });
    expect(countTags(svg, 'foreignObject')).toBeGreaterThan(0);
    const text = textContentOf(svg);
    expect(text).toContain('Styled');
    expect(text).toContain('Plain');
    // The styled cell's run still carries its bold weight and red color.
    expect(svg).toContain('font-weight:700');
    expect(svg).toMatch(/color:#[Cc][Cc]0+0+/);
  });
});
