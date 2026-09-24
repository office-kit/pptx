import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import * as pptx from '../src/api/index.ts';

it('patches selected cell paragraphs without flattening runs and survives export', async () => {
  const presentation = await pptx.loadPresentation(
    await readFile(new URL('./fixtures/minimal/one-text-slide.pptx', import.meta.url)),
  );
  const slide = pptx.getSlides(presentation)[0]!;
  const table = pptx.addSlideTable(slide, {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(4),
    h: pptx.inches(2),
    rows: [['', 'other']],
  });
  const cell = pptx.getTableCell(table, 0, 0);
  pptx.setTableCellParagraphs(cell, [
    { align: 'center', runs: [{ text: 'First', format: { bold: true } }] },
    { align: 'right', runs: [{ text: 'Second', format: { italic: true } }] },
  ]);
  pptx.setTableCellTextRangeParagraphSettings(cell, 6, 12, {
    leftEmu: 182880,
    rightEmu: 91440,
    firstLineEmu: -91440,
    beforePts: 6,
    afterPts: 8,
    lineSpacing: { kind: 'pts', value: 30 },
  });
  const before = pptx.getSlideXmlString(slide);
  expect(() => pptx.shiftTableCellTextRangeLevel(cell, 6, 12, 2 as 1)).toThrow();
  expect(() => pptx.shiftTableCellTextRangeLevel(cell, -1, 12, 1)).toThrow();
  expect(pptx.getSlideXmlString(slide)).toBe(before);
  expect(() =>
    pptx.setTableCellTextRangeParagraphSettings(cell, 0, 5, { align: 'left', afterPts: -1 }),
  ).toThrow();
  expect(pptx.getSlideXmlString(slide)).toBe(before);
  for (let i = 0; i < 10; i++) pptx.shiftTableCellTextRangeLevel(cell, 6, 12, 1);
  expect(pptx.getTableCellParagraphs(cell)[1]!.properties?.level).toBe(8);
  pptx.shiftTableCellTextRangeLevel(cell, 0, 12, -1);
  const loaded = await pptx.loadPresentation(await pptx.savePresentation(presentation));
  const resultTable = pptx.getSlideTables(pptx.getSlides(loaded)[0]!)[0]!;
  const paragraphs = pptx.getTableCellParagraphs(pptx.getTableCell(resultTable, 0, 0));
  expect(paragraphs[0]!.align).toBe('center');
  expect(paragraphs[0]!.properties?.level).toBe(0);
  expect(paragraphs[0]!.properties?.spcBefPts).toBeUndefined();
  expect(paragraphs[0]!.elements[0]!.format?.bold).toBe(true);
  expect(paragraphs[1]!.align).toBe('right');
  expect(paragraphs[1]!.properties).toMatchObject({
    level: 7,
    marL: 182880,
    marR: 91440,
    indent: -91440,
    spcBefPts: 6,
    spcAftPts: 8,
    lineSpacing: { kind: 'pts', value: 30 },
  });
  expect(paragraphs[1]!.elements[0]!.format?.italic).toBe(true);
  expect(pptx.getTableCellText(pptx.getTableCell(resultTable, 0, 1))).toBe('other');
});
