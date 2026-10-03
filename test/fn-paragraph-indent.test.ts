import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  addSlideTable,
  createPresentation,
  getParagraphIndent,
  getSlideShapes,
  getSlides,
  getTableCell,
  inches,
  loadPresentation,
  savePresentation,
  setParagraphIndent,
  setParagraphAlignment,
  getParagraphAlignment,
} from '../src/api/index.ts';

const makeBox = () => {
  const pres = createPresentation();
  const box = addSlideTextBox(addBlankSlide(pres), {
    x: inches(1),
    y: inches(1),
    w: inches(6),
    h: inches(2),
    text: 'first\nsecond',
  });
  return { pres, box };
};

describe('setParagraphIndent', () => {
  it('preserves omitted sides, other paragraphs and paragraph formatting across export', async () => {
    const { pres, box } = makeBox();
    setParagraphAlignment(box, 0, 'right');
    setParagraphIndent(box, 0, { leftEmu: 360000, rightEmu: 720000, firstLineEmu: -180000 });
    setParagraphIndent(box, 0, { leftEmu: 450000.4 });
    const loaded = getSlideShapes(
      getSlides(await loadPresentation(await savePresentation(pres)))[0]!,
    )[0]!;
    expect(getParagraphIndent(loaded, 0)).toEqual({
      leftEmu: 450000,
      rightEmu: 720000,
      firstLineEmu: -180000,
    });
    expect(getParagraphAlignment(loaded, 0)).toBe('r');
    expect(getParagraphIndent(loaded, 1)).toEqual({
      leftEmu: null,
      rightEmu: null,
      firstLineEmu: null,
    });
    setParagraphIndent(loaded, 0, { rightEmu: null, firstLineEmu: null });
    expect(getParagraphIndent(loaded, 0)).toEqual({
      leftEmu: 450000,
      rightEmu: null,
      firstLineEmu: null,
    });
  });

  it.each([NaN, Infinity, -1, 51206401])('rejects invalid margins atomically: %s', (value) => {
    const { box } = makeBox();
    setParagraphIndent(box, 0, { leftEmu: 360000 });
    expect(() => setParagraphIndent(box, 0, { leftEmu: 720000, rightEmu: value })).toThrow(
      RangeError,
    );
    expect(getParagraphIndent(box, 0)).toEqual({
      leftEmu: 360000,
      rightEmu: null,
      firstLineEmu: null,
    });
  });

  it('accepts signed first-line bounds and rejects values outside them', () => {
    const { box } = makeBox();
    for (const value of [-51206400, 51206400]) {
      setParagraphIndent(box, 0, { firstLineEmu: value });
      expect(getParagraphIndent(box, 0).firstLineEmu).toBe(value);
    }
    for (const value of [-51206401, 51206401, NaN]) {
      expect(() => setParagraphIndent(box, 0, { firstLineEmu: value })).toThrow(RangeError);
    }
  });

  it('persists table-cell indentation', async () => {
    const pres = createPresentation();
    const table = addSlideTable(addBlankSlide(pres), {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(2),
      rows: [['left', 'right']],
    });
    setParagraphIndent(getTableCell(table, 0, 0)!, 0, { leftEmu: 360000, firstLineEmu: -180000 });
    const loaded = getSlideShapes(
      getSlides(await loadPresentation(await savePresentation(pres)))[0]!,
    )[0]!;
    expect(getParagraphIndent(getTableCell(loaded, 0, 0)!, 0)).toEqual({
      leftEmu: 360000,
      rightEmu: null,
      firstLineEmu: -180000,
    });
    expect(getParagraphIndent(getTableCell(loaded, 0, 1)!, 0)).toEqual({
      leftEmu: 0,
      rightEmu: null,
      firstLineEmu: 0,
    });
  });
});
