import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import {
  addSlideTable,
  removeShape,
  getSlides,
  getSlideShapes,
  getTableCells,
  getTableCellText,
  getTableCellParagraphs,
  isTableShape,
  inches,
  loadPresentation,
  savePresentation,
  setParagraphAlignment,
  setParagraphBullet,
  setParagraphLevel,
  setParagraphLineSpacing,
  setParagraphSpacing,
  getParagraphPropertiesEffective,
  setTableCellTextFormat,
} from '../src/api/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

it('edits cell paragraphs without replacing rich text and round-trips their visible formatting', async () => {
  const pres = await loadPresentation(
    await readFile(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url)),
  );
  const slide = getSlides(pres)[0]!;
  getSlideShapes(slide).forEach(removeShape);
  const table = addSlideTable(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(6),
    h: inches(3),
    rows: [['日本語\nEnglish', 'Untouched']],
  });
  const cell = getTableCells(table)[0]![0]!;
  setTableCellTextFormat(cell, { bold: true, color: '#CC0000' });
  setParagraphBullet(cell, 1, 'bullet');
  setParagraphLevel(cell, 1, 1);
  setParagraphAlignment(cell, 1, 'right');
  setParagraphLineSpacing(cell, 1, { kind: 'pct', value: 1.5 });
  setParagraphSpacing(cell, 1, { beforePts: 8, afterPts: 4 });
  const expected = getParagraphPropertiesEffective(pres, cell, 1);
  expect(expected).toMatchObject({
    align: 'right',
    level: 1,
    bullet: 'bullet',
    lineSpacing: { kind: 'pct', value: 1.5 },
    spcBefPts: 8,
    spcAftPts: 4,
  });
  expect(getParagraphPropertiesEffective(pres, cell, 0).bullet).toBe('none');
  expect(getTableCellText(cell)).toBe('日本語\nEnglish');
  expect(getTableCellParagraphs(cell)[1]!.elements[0]).toMatchObject({
    text: 'English',
    format: { bold: true, color: '#CC0000' },
  });
  const loaded = await loadPresentation(await savePresentation(pres));
  const loadedSlide = getSlides(loaded)[0]!;
  const loadedTable = getSlideShapes(loadedSlide).find(isTableShape)!;
  expect(getParagraphPropertiesEffective(loaded, getTableCells(loadedTable)[0]![0]!, 1)).toEqual(
    expected,
  );
  expect(getTableCellText(getTableCells(loadedTable)[0]![1]!)).toBe('Untouched');
  for (const textLayout of ['svg', 'foreignObject'] as const) {
    const svg = renderSlideToSvg(loaded, loadedSlide, { textLayout });
    expect(svg).toContain('◦');
    if (textLayout === 'foreignObject') {
      expect(svg).toContain('line-height:1.500');
      expect(svg).toContain('margin-top:10.67px');
    }
    expect(svg).toContain('English');
  }
  const loadedCell = getTableCells(loadedTable)[0]![0]!;
  setParagraphBullet(loadedCell, 0, 'number');
  setParagraphBullet(loadedCell, 1, 'number');
  setParagraphLevel(loadedCell, 1, 0);
  for (const textLayout of ['svg', 'foreignObject'] as const) {
    const svg = renderSlideToSvg(loaded, loadedSlide, { textLayout });
    expect(svg).toContain('1.');
    expect(svg).toContain('2.');
  }
});
