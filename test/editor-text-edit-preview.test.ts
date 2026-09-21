import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  addSlideTable,
  createPresentation,
  inches,
  getShapeText,
  getShapeParagraphElements,
  getTableCells,
  getTableCellText,
  getTableCellParagraphs,
  setShapeParagraphs,
  setTableCellTextFormat,
  getSlideShapes,
} from '../src/api/index.ts';
import { projectTextEdits } from '../site/src/lib/editor/core/text-edit-preview.ts';

describe('pending text formatting preview', () => {
  it('preserves sequential edit inheritance without mutating the source shape', () => {
    const slide = addBlankSlide(createPresentation());
    const shape = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(2),
      text: '',
    });
    setShapeParagraphs(shape, [
      { runs: [{ text: 'English', format: { bold: true } }] },
      { runs: [{ text: '日本語', format: { italic: true } }] },
    ]);
    const projected = projectTextEdits(shape, [
      { start: 0, end: 0, text: 'Prefix\n' },
      { start: 15, end: 18, text: '日本語です' },
    ]);
    expect(getShapeText(projected)).toBe('Prefix\nEnglish\n日本語です');
    expect(getShapeParagraphElements(projected, 1)[0]!.format?.bold).toBe(true);
    expect(getShapeParagraphElements(projected, 2)[0]!.format?.italic).toBe(true);
    expect(getShapeText(shape)).toBe('English\n日本語');
    expect(getSlideShapes(slide)).toHaveLength(1);
    expect(projectTextEdits(shape, [])).toBe(shape);
  });
  it('projects only the edited table cell and leaves the live table unchanged', () => {
    const slide = addBlankSlide(createPresentation());
    const shape = addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(2),
      rows: [['English', 'Other']],
    });
    setTableCellTextFormat(getTableCells(shape)[0]![0]!, { underline: true });
    const projected = projectTextEdits(shape, [{ start: 0, end: 0, text: '日本語\n' }], {
      row: 0,
      col: 0,
    });
    const cells = getTableCells(projected);
    expect(getTableCellText(cells[0]![0]!)).toBe('日本語\nEnglish');
    expect(getTableCellParagraphs(cells[0]![0]!)[1]!.elements[0]!.format?.underline).toBe(true);
    expect(getTableCellText(cells[0]![1]!)).toBe('Other');
    expect(getTableCellText(getTableCells(shape)[0]![0]!)).toBe('English');
  });
});
