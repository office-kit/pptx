import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import * as pptx from '../src/api/index.ts';

const fixture = async () =>
  pptx.loadPresentation(
    await readFile(new URL('./fixtures/minimal/one-text-slide.pptx', import.meta.url)),
  );

it('persists cell text links, preserves formatting and leaves neighboring cells and object links intact', async () => {
  const p = await fixture();
  const slide = pptx.getSlides(p)[0]!;
  const table = pptx.addSlideTable(slide, {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(4),
    h: pptx.inches(2),
    rows: [['', 'neighbor']],
  });
  const cell = pptx.getTableCell(table, 0, 0);
  pptx.setTableCellParagraphs(cell, [
    { runs: [{ text: 'A😀B', format: { bold: true } }] },
    { runs: [{ text: '日本語', format: { italic: true } }] },
  ]);
  pptx.setShapeClickAction(table, { kind: 'nextSlide' }, 'Object');
  pptx.setTableCellTextRangeClickAction(
    cell,
    1,
    7,
    { kind: 'url', url: 'https://example.com/?a=1&b=2' },
    '資料 <詳細>',
  );
  const expected = [
    {
      start: 1,
      end: 4,
      action: { kind: 'url', url: 'https://example.com/?a=1&b=2' },
      tooltip: '資料 <詳細>',
    },
    {
      start: 5,
      end: 7,
      action: { kind: 'url', url: 'https://example.com/?a=1&b=2' },
      tooltip: '資料 <詳細>',
    },
  ];
  const check = (table: pptx.SlideShapeData) => {
    const cell = pptx.getTableCell(table, 0, 0);
    expect(pptx.getTableCellTextRangeClickActions(cell)).toEqual(expected);
    expect(pptx.getTableCellText(cell)).toBe('A😀B\n日本語');
    expect(pptx.getTableCellParagraphs(cell)[0]!.elements.every((e) => e.format?.bold)).toBe(true);
    expect(pptx.getTableCellParagraphs(cell)[1]!.elements.every((e) => e.format?.italic)).toBe(
      true,
    );
    expect(pptx.getTableCellText(pptx.getTableCell(table, 0, 1))).toBe('neighbor');
    expect(pptx.getTableCellTextRangeClickActions(pptx.getTableCell(table, 0, 1))).toEqual([]);
    expect(pptx.getShapeClickAction(table)).toEqual({ kind: 'nextSlide' });
  };
  check(table);
  const loaded = await pptx.loadPresentation(await pptx.savePresentation(p));
  check(pptx.getSlideTables(pptx.getSlides(loaded)[0]!)[0]!);
  pptx.setTableCellTextRangeClickAction(cell, 1, 3, null);
  expect(pptx.getTableCellTextRangeClickActions(cell)).toEqual([
    { ...expected[0], start: 3 },
    expected[1],
  ]);
});

it('round-trips cell slide destinations and rejects invalid ranges or foreign slides atomically', async () => {
  const p = await fixture();
  const slide = pptx.getSlides(p)[0]!;
  const destination = pptx.addBlankSlide(p);
  const table = pptx.addSlideTable(slide, {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(4),
    h: pptx.inches(2),
    rows: [['A😀B']],
  });
  const cell = pptx.getTableCell(table, 0, 0);
  pptx.setTableCellTextRangeClickAction(cell, 1, 3, { kind: 'slide', slide: destination }, 'Jump');
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  const link = pptx.getTableCellTextRangeClickActions(
    pptx.getTableCell(pptx.getSlideTables(pptx.getSlides(saved)[0]!)[0]!, 0, 0),
  )[0]!;
  expect(link.action.kind).toBe('slide');
  if (link.action.kind === 'slide')
    expect(pptx.getSlidePartName(link.action.slide)).toBe(pptx.getSlidePartName(destination));
  const before = pptx.getSlideXmlString(slide);
  for (const [start, end] of [
    [-1, 0],
    [1, 2],
    [3, 1],
    [0, 99],
  ]) {
    expect(() => pptx.setTableCellTextRangeClickAction(cell, start!, end!, null)).toThrow();
    expect(pptx.getSlideXmlString(slide)).toBe(before);
  }
  const foreign = await fixture();
  expect(() =>
    pptx.setTableCellTextRangeClickAction(cell, 0, 1, {
      kind: 'slide',
      slide: pptx.getSlides(foreign)[0]!,
    }),
  ).toThrow();
  expect(pptx.getSlideXmlString(slide)).toBe(before);
  pptx.setTableCellTextRangeClickAction(cell, 0, 0, {
    kind: 'url',
    url: 'https://example.com/unused',
  });
  expect(pptx.getSlideXmlString(slide)).toBe(before);
  for (const kind of ['nextSlide', 'prevSlide', 'firstSlide', 'lastSlide'] as const) {
    pptx.setTableCellTextRangeClickAction(cell, 0, 1, { kind });
    expect(pptx.getTableCellTextRangeClickActions(cell)[0]!.action).toEqual({ kind });
  }
});
