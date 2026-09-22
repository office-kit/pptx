import { readFile } from 'node:fs/promises';
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
  getParagraphPropertiesEffective,
  getParagraphBullet,
  loadPresentation,
  findSlideLayout,
  findSlidePlaceholder,
  addSlide,
  setShapeText,
  setParagraphLevel,
  setParagraphLineSpacing,
} from '../src/api/index.ts';
import { projectTextEdits } from '../site/src/lib/editor/core/text-edit-preview.ts';

describe('pending text formatting preview', () => {
  it('preserves sequential edit inheritance without mutating the source shape', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const shape = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(2),
      text: '',
    });
    setShapeParagraphs(shape, [
      { runs: [{ text: 'English', format: { bold: true } }] },
      {
        align: 'center',
        runs: [{ text: '日本語', format: { italic: true } }],
      },
    ]);
    setParagraphLevel(shape, 1, 3);
    setParagraphLineSpacing(shape, 1, { kind: 'pct', value: 1.5 });
    const projected = projectTextEdits(
      shape,
      [
        { start: 0, end: 0, text: 'Prefix\n' },
        { start: 15, end: 18, text: '日本語です' },
      ],
      undefined,
      pres,
    );
    expect(getParagraphPropertiesEffective(pres, projected, 2)).toMatchObject({
      align: 'center',
      level: 3,
      lineSpacing: { kind: 'pct', value: 1.5 },
    });
    expect(getShapeText(projected)).toBe('Prefix\nEnglish\n日本語です');
    expect(getShapeParagraphElements(projected, 1)[0]!.format?.bold).toBe(true);
    expect(getShapeParagraphElements(projected, 2)[0]!.format?.italic).toBe(true);
    expect(getShapeText(shape)).toBe('English\n日本語');
    expect(getSlideShapes(slide)).toHaveLength(1);
    expect(projectTextEdits(shape, [])).toBe(shape);
  });
  it('retains master paragraph defaults while projecting placeholder edits', async () => {
    const pres = await loadPresentation(
      await readFile(new URL('./fixtures/minimal/blank.pptx', import.meta.url)),
    );
    const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Title and Content')! });
    const shape = findSlidePlaceholder(slide, 'body')!;
    setShapeText(shape, 'English');
    const before = getParagraphPropertiesEffective(pres, shape, 0);
    expect(before.bullet).toBe('bullet');
    const projected = projectTextEdits(
      shape,
      [{ start: 0, end: 0, text: '日本語\n' }],
      undefined,
      pres,
    );
    for (const index of [0, 1]) {
      const props = getParagraphPropertiesEffective(pres, projected, index);
      expect(props.bullet).toBe(before.bullet);
      expect(props.align).toBe(before.align);
      expect(props.lineSpacing).toEqual(before.lineSpacing);
      expect(props.spcBefPts).toBe(before.spcBefPts);
      expect(props.spcAftPts).toBe(before.spcAftPts);
    }
    expect(getParagraphBullet(shape, 0)).toBeNull();
    expect(getShapeText(shape)).toBe('English');
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
