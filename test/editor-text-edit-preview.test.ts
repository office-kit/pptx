import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  addSlideTable,
  createPresentation,
  inches,
  getShapeText,
  getShapeRunFormatEffective,
  getShapeXmlString,
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
import { copyTextRange } from '../site/src/lib/editor/core/text-clipboard.ts';
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
  it('resolves inherited display styles after detached edits without baking them into text', async () => {
    const pres = await loadPresentation(
      await readFile(new URL('./fixtures/minimal/blank.pptx', import.meta.url)),
    );
    const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Title and Content')! });
    const shape = findSlidePlaceholder(slide, 'body')!;
    setShapeText(shape, 'English');
    const xml = getShapeXmlString(shape);
    const inherited = getShapeRunFormatEffective(pres, shape, 0, 0);
    expect(inherited.size).toBeGreaterThan(18);
    const projected = projectTextEdits(
      shape,
      [{ start: 7, end: 7, text: '\n日本語', typing: { format: { italic: true }, reset: true } }],
      undefined,
      pres,
    );
    const literal = copyTextRange(projected, 0, 11);
    expect(literal.formats.every((span) => span.format.size === undefined)).toBe(true);
    const display = copyTextRange(projected, 0, 11, undefined, (paragraph, run) =>
      getShapeRunFormatEffective(pres, projected, paragraph, run, { inheritanceSource: shape }),
    );
    expect(display.text).toBe('English\n日本語');
    const originalParagraph = getParagraphPropertiesEffective(pres, shape, 0);
    const previewParagraph = getParagraphPropertiesEffective(pres, projected, 1, {
      inheritanceSource: shape,
    });
    expect(originalParagraph.marL).toBeGreaterThan(0);
    expect(previewParagraph.marL).toBe(originalParagraph.marL);
    expect(previewParagraph.indent).toBe(originalParagraph.indent);
    expect(display.formats[0]!.format).toMatchObject(inherited);
    expect(display.formats.at(-1)!.format).toMatchObject({ ...inherited, italic: true });
    expect(getShapeXmlString(shape)).toBe(xml);
    expect(copyTextRange(projected, 0, 11)).toEqual(literal);
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
  it('applies typing snapshots only to inserted characters, including composition replacements', () => {
    const slide = addBlankSlide(createPresentation());
    const shape = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(2),
      text: 'AB',
    });
    const projected = projectTextEdits(shape, [
      { start: 1, end: 1, text: 'に', typing: { format: { bold: true, size: 32 }, reset: false } },
      {
        start: 1,
        end: 2,
        text: '日本語',
        typing: { format: { bold: true, size: 32 }, reset: false },
      },
      { start: 4, end: 4, text: '😀', typing: { format: {}, reset: true } },
    ]);
    const runs = getShapeParagraphElements(projected, 0).filter((r) => r.kind !== 'br');
    expect(runs.map((r) => r.text).join('')).toBe('A日本語😀B');
    expect(runs.find((r) => r.text.includes('日本語'))?.format).toMatchObject({
      bold: true,
      size: 32,
    });
    expect(runs.find((r) => r.text.includes('😀'))?.format?.bold).toBeUndefined();
    expect(runs.find((r) => r.text.includes('B'))?.format?.bold).toBeUndefined();
    expect(getShapeText(shape)).toBe('AB');
  });
});
