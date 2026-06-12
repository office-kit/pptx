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
  setTableCellAlignment,
  setTableCellTextFormat,
} from '../src/api/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';
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
