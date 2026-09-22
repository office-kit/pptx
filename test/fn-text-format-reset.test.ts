import { describe, expect, it } from 'vitest';
import {
  createPresentation,
  addBlankSlide,
  addSlideTextBox,
  addSlideTable,
  inches,
  setShapeTextFormat,
  setTableCellTextFormat,
  getTableCells,
  getTableCellParagraphs,
  getShapeParagraphElements,
  getShapeText,
  setShapeHyperlink,
  getShapeRunHyperlink,
  setParagraphAlignment,
  getParagraphAlignment,
  getSlides,
  getSlideShapes,
  savePresentation,
  loadPresentation,
} from '../src/api/index.ts';
import { applyFormatToAllRuns } from '../src/internal/drawingml/text-format.ts';
import { parseXml, serializeFragment } from '../src/internal/xml/index.ts';

const decorated = {
  font: 'Arial',
  fontEastAsian: 'Meiryo',
  fontComplexScript: 'Arial',
  size: 32,
  bold: true,
  italic: true,
  underline: true,
  strike: true,
  color: '#FF0000',
  highlight: '#FFFF00',
  baseline: 0.3,
  spc: 100,
  kern: 1200,
  cap: 'all' as const,
};

describe('reset text appearance', () => {
  it('clears selected characters, preserving links, surrounding runs and paragraph alignment through reload', async () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const shape = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(2),
      text: 'A日本😀Z',
    });
    setShapeTextFormat(shape, decorated);
    setShapeHyperlink(shape, 'https://example.com');
    setParagraphAlignment(shape, 0, 'center');
    setShapeTextFormat(shape, {}, { reset: true, range: { start: 1, end: 5 } });
    const loaded = await loadPresentation(await savePresentation(pres));
    const result = getSlideShapes(getSlides(loaded)[0]!)[0]!;
    expect(getShapeText(result)).toBe('A日本😀Z');
    const elements = getShapeParagraphElements(result, 0);
    expect(elements.map((e) => (e.kind === 'br' ? '\n' : e.text))).toEqual(['A', '日本😀', 'Z']);
    expect(elements[0]!.format).toMatchObject(decorated);
    expect(elements[1]!.format ?? {}).toEqual({});
    expect(elements[2]!.format).toMatchObject(decorated);
    expect(getShapeRunHyperlink(result, 0, 1)).toBe('https://example.com');
    expect(getParagraphAlignment(result, 0)).toBe('ctr');
  });
  it('resets a table-cell range and supports replacing the whole cell format', async () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const shape = addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(2),
      rows: [['A日本Z', 'Other']],
    });
    const [cell, other] = getTableCells(shape)[0]!;
    setTableCellTextFormat(cell!, decorated);
    setTableCellTextFormat(other!, { bold: true });
    setTableCellTextFormat(cell!, {}, { reset: true, range: { start: 1, end: 3 } });
    expect(getTableCellParagraphs(cell!)[0]!.elements[1]!.format ?? {}).toEqual({});
    expect(getTableCellParagraphs(cell!)[0]!.elements[0]!.format?.bold).toBe(true);
    setTableCellTextFormat(cell!, { italic: true }, { reset: true });
    const loaded = await loadPresentation(await savePresentation(pres));
    const cells = getTableCells(getSlideShapes(getSlides(loaded)[0]!)[0]!)[0]!;
    for (const element of getTableCellParagraphs(cells[0]!)[0]!.elements)
      expect(element.format).toEqual({ italic: true });
    expect(getTableCellParagraphs(cells[1]!)[0]!.elements[0]!.format?.bold).toBe(true);
  });
  it('clears fields, breaks and defaults while retaining language, links and unknown XML', () => {
    const body = parseXml(
      '<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:custom"><a:lstStyle><a:lvl1pPr><a:defRPr sz="2400"/></a:lvl1pPr></a:lstStyle><a:p><a:pPr algn="ctr"><a:defRPr b="1"/></a:pPr><a:fld id="field"><a:rPr i="1" lang="ja-JP" x:b="keep"><a:solidFill/><a:hlinkClick/><a:extLst/><x:solidFill/></a:rPr><a:t>1</a:t></a:fld><a:br><a:rPr u="sng"/></a:br><a:endParaRPr sz="3000"/></a:p></a:txBody>',
    ).root;
    applyFormatToAllRuns(body, {}, 'reset test', true);
    const xml = serializeFragment(body);
    expect(xml).not.toMatch(/\s(?:sz|b|i|u)="/);
    expect(xml).toContain('lang="ja-JP"');
    expect(xml).toContain('x:b="keep"');
    expect(xml).toContain('algn="ctr"');
    expect(xml).toContain('a:hlinkClick');
    expect(xml).toContain('a:extLst');
    expect(xml).toContain('x:solidFill');
    expect(xml).not.toContain('<a:solidFill');
  });
  it('does not clear appearance when replacement formatting is invalid', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const shape = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(2),
      text: 'Test',
    });
    setShapeTextFormat(shape, { bold: true });
    expect(() => setShapeTextFormat(shape, { size: -1 }, { reset: true })).toThrow();
    expect(getShapeParagraphElements(shape, 0)[0]!.format?.bold).toBe(true);
  });
});
