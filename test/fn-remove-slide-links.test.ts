import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import {
  addSlide,
  addSlideTable,
  getSlideLayouts,
  getSlideShapes,
  getSlides,
  getShapeClickAction,
  getShapeRunClickAction,
  getShapeParagraphElements,
  getTableCell,
  getTableCellParagraphs,
  getSlideXmlString,
  getSlidePartName,
  inches,
  loadPresentation,
  removeSlide,
  savePresentation,
  setShapeClickAction,
  setShapeTextFormat,
  setTableCellClickAction,
} from '../src/api/index.ts';

const fixture = async () =>
  loadPresentation(await readFile(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url)));

it('clears deleted slide links from objects, text and cells before part names are reused', async () => {
  const pres = await fixture();
  const [source, target] = getSlides(pres);
  const shape = getSlideShapes(source!)[0]!;
  const table = addSlideTable(source!, {
    x: inches(1),
    y: inches(2),
    w: inches(6),
    h: inches(1),
    rows: [['日本語', 'Keep']],
  });
  const cell = getTableCell(table, 0, 0);
  const neighbor = getTableCell(table, 0, 1);
  const action = { kind: 'slide' as const, slide: target! };
  setShapeClickAction(shape, action, { tooltip: 'Object destination' });
  setShapeClickAction(shape, action, { range: { start: 0, end: 1 }, tooltip: 'Text destination' });
  setShapeTextFormat(shape, { bold: true }, { range: { start: 0, end: 1 } });
  setTableCellClickAction(cell, action, { tooltip: '表の資料' });
  setTableCellClickAction(
    neighbor,
    { kind: 'url', url: 'https://example.com/keep' },
    { tooltip: 'Keep this' },
  );
  const check = () => {
    expect(getShapeClickAction(shape)).toBeNull();
    expect(getShapeRunClickAction(shape, 0, 0)).toBeNull();
    expect(getShapeParagraphElements(shape, 0)[0]!.format?.bold).toBe(true);
    expect(getTableCellParagraphs(cell)[0]!.elements[0]!.clickAction).toBeUndefined();
    expect(getTableCellParagraphs(cell)[0]!.elements[0]!.tooltip).toBeUndefined();
    expect(getTableCellParagraphs(neighbor)[0]!.elements[0]!.clickAction).toEqual({
      kind: 'url',
      url: 'https://example.com/keep',
    });
    expect(getTableCellParagraphs(neighbor)[0]!.elements[0]!.tooltip).toBe('Keep this');
    expect(getSlideXmlString(source!)).not.toContain('ppaction://hlinksldjump');
  };
  const oldPart = getSlidePartName(target!);
  removeSlide(pres, target!);
  check();
  const replacement = addSlide(pres, { layout: getSlideLayouts(pres)[0]! });
  expect(getSlidePartName(replacement)).toBe(oldPart);
  check();
  // Later edits through retained handles must not bring deleted links back.
  setShapeTextFormat(shape, { italic: true });
  const loaded = await loadPresentation(await savePresentation(pres));
  expect(getSlideXmlString(getSlides(loaded)[0]!)).not.toContain('ppaction://hlinksldjump');
  expect(getShapeClickAction(getSlideShapes(getSlides(loaded)[0]!)[0]!)).toBeNull();
});

it('rejects a foreign slide handle with the same part name without changing the deck', async () => {
  const pres = await fixture();
  const foreign = await fixture();
  const before = await savePresentation(pres);
  expect(() => removeSlide(pres, getSlides(foreign)[1]!)).toThrow('must belong');
  expect(await savePresentation(pres)).toEqual(before);
});
