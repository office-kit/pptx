import { expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  addSlideTable,
  createPresentation,
  getParagraphLevel,
  getParagraphIndent,
  getShapeText,
  getShapeXmlString,
  getSlideShapes,
  getSlides,
  getTableCells,
  inches,
  loadPresentation,
  savePresentation,
  setParagraphLevel,
  setShapeParagraphs,
} from '../src/api/index.ts';

function fixture() {
  const pres = createPresentation();
  const shape = addSlideTextBox(addBlankSlide(pres), {
    x: inches(1),
    y: inches(1),
    w: inches(6),
    h: inches(4),
    text: '',
  });
  setShapeParagraphs(shape, [
    { runs: [{ text: '日本語', format: { bold: true } }] },
    { runs: [{ text: 'Second\nline' }] },
    { runs: [{ text: 'Last' }] },
  ]);
  setParagraphLevel(shape, 1, 2);
  setParagraphLevel(shape, 2, 8);
  return { pres, shape };
}

it('changes selected paragraph levels relative to each other and persists without changing text', async () => {
  const { pres, shape } = fixture();
  const text = getShapeText(shape);
  setParagraphLevel(shape, { start: 0, end: text.length }, { offset: 1 });
  expect([0, 1, 2].map((index) => getParagraphLevel(shape, index))).toEqual([1, 3, 8]);
  expect(getShapeText(shape)).toBe(text);
  expect(getParagraphLevel(shape, { start: 0, end: text.length })).toEqual([1, 3, 8]);
  expect(getShapeXmlString(shape)).toContain('b="1"');
  const reloaded = getSlideShapes(
    getSlides(await loadPresentation(await savePresentation(pres)))[0]!,
  )[0]!;
  expect([0, 1, 2].map((index) => getParagraphLevel(reloaded, index))).toEqual([1, 3, 8]);
  setParagraphLevel(reloaded, { start: 0, end: text.length }, { offset: -4 });
  expect([0, 1, 2].map((index) => getParagraphLevel(reloaded, index))).toEqual([0, 0, 4]);
});

it('keeps newlines inside a run in one paragraph and treats the range end as exclusive', () => {
  const { shape } = fixture();
  setParagraphLevel(shape, { start: 0, end: 4 }, 4);
  expect([0, 1, 2].map((index) => getParagraphLevel(shape, index))).toEqual([4, 2, 8]);
  setParagraphLevel(shape, { start: 11, end: 11 }, { offset: 1 });
  expect([0, 1, 2].map((index) => getParagraphLevel(shape, index))).toEqual([4, 3, 8]);
  expect(getParagraphIndent(shape, 1)).toEqual({
    leftEmu: null,
    rightEmu: null,
    firstLineEmu: null,
  });
});

it('validates the entire range and offset before changing XML', () => {
  const { shape } = fixture();
  const before = getShapeXmlString(shape);
  for (const range of [
    { start: -1, end: 2 },
    { start: 2, end: 1 },
    { start: 0, end: 200 },
    { start: 0.5, end: 2 },
  ]) {
    expect(() => setParagraphLevel(shape, range, { offset: 1 })).toThrow(RangeError);
    expect(getShapeXmlString(shape)).toBe(before);
  }
  for (const offset of [NaN, Infinity, 0.5]) {
    expect(() => setParagraphLevel(shape, { start: 0, end: 2 }, { offset })).toThrow(RangeError);
    expect(getShapeXmlString(shape)).toBe(before);
  }
});

it('includes empty final paragraphs and rejects split surrogate boundaries', () => {
  const { shape } = fixture();
  setShapeParagraphs(shape, [{ runs: [{ text: '😀' }] }, { runs: [] }]);
  setParagraphLevel(shape, { start: 3, end: 3 }, 2);
  expect(getParagraphLevel(shape, { start: 0, end: 3 })).toEqual([0]);
  expect(getParagraphLevel(shape, { start: 3, end: 3 })).toEqual([2]);
  const before = getShapeXmlString(shape);
  expect(() => setParagraphLevel(shape, { start: 1, end: 2 }, 3)).toThrow(RangeError);
  expect(getShapeXmlString(shape)).toBe(before);
});

it('updates a table cell range without changing adjacent cells and persists levels', async () => {
  const pres = createPresentation();
  const table = addSlideTable(addBlankSlide(pres), {
    x: inches(1),
    y: inches(1),
    w: inches(6),
    h: inches(2),
    rows: [['First\nSecond', 'Neighbor']],
  });
  const [cell, neighbor] = getTableCells(table)[0]!;
  setParagraphLevel(cell!, { start: 0, end: 12 }, { offset: 2 });
  expect(getParagraphLevel(cell!, { start: 0, end: 12 })).toEqual([2, 2]);
  expect(getParagraphLevel(neighbor!, 0)).toBe(0);
  const loaded = await loadPresentation(await savePresentation(pres));
  const restored = getTableCells(getSlideShapes(getSlides(loaded)[0]!)[0]!)[0]!;
  expect(getParagraphLevel(restored[0]!, { start: 0, end: 12 })).toEqual([2, 2]);
  expect(getParagraphLevel(restored[1]!, 0)).toBe(0);
});
