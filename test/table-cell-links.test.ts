import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import {
  addSlideTable,
  getSlides,
  getSlideTables,
  getSlideIndex,
  getTableCell,
  getTableCellParagraphs,
  getTableCellText,
  getShapeClickAction,
  inches,
  loadPresentation,
  savePresentation,
  setTableCellClickAction,
  setTableCellTextFormat,
  type ShapeClickAction,
} from '../src/api/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

const fixture = async () => {
  const pres = await loadPresentation(
    await readFile(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url)),
  );
  const slide = getSlides(pres)[0]!;
  const table = addSlideTable(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(7),
    h: inches(2),
    rows: [['Before 日本語 after', 'Untouched']],
  });
  const cell = getTableCell(table, 0, 0);
  setTableCellTextFormat(cell, { bold: true });
  return { pres, slide, table, cell };
};

it('round-trips every cell click action and description while preserving surrounding text and adjacent cells', async () => {
  const { pres, table, cell } = await fixture();
  const actions: ShapeClickAction[] = [
    { kind: 'url', url: 'https://example.com/?a=1&b=2' },
    { kind: 'slide', slide: getSlides(pres)[1]! },
    { kind: 'nextSlide' },
    { kind: 'prevSlide' },
    { kind: 'firstSlide' },
    { kind: 'lastSlide' },
  ];
  const range = { start: 7, end: 10 };
  const neighbor = getTableCellParagraphs(getTableCell(table, 0, 1));
  for (const action of actions) {
    setTableCellClickAction(cell, action, { range, tooltip: '参考 "Link" & <資料>' });
    const loaded = await loadPresentation(await savePresentation(pres));
    const loadedTable = getSlideTables(getSlides(loaded)[0]!)[0]!;
    const loadedCell = getTableCell(loadedTable, 0, 0);
    expect(getTableCellText(loadedCell)).toBe('Before 日本語 after');
    const elements = getTableCellParagraphs(loadedCell)[0]!.elements;
    expect(elements).toHaveLength(3);
    expect(elements.map((e) => (e.kind === 'br' ? '\n' : e.text))).toEqual([
      'Before ',
      '日本語',
      ' after',
    ]);
    expect(elements.every((e) => e.format?.bold)).toBe(true);
    expect(elements[0]!.clickAction).toBeUndefined();
    expect(elements[2]!.clickAction).toBeUndefined();
    expect(elements[1]!.tooltip).toBe('参考 "Link" & <資料>');
    const read = elements[1]!.clickAction!;
    if (read.kind === 'slide') expect(getSlideIndex(loaded, read.slide)).toBe(1);
    else expect(read).toEqual(action);
    expect(getTableCellParagraphs(getTableCell(loadedTable, 0, 1))).toEqual(neighbor);
    expect(getShapeClickAction(loadedTable)).toBeNull();
  }
  setTableCellClickAction(cell, { kind: 'nextSlide' }, { range });
  expect(getTableCellParagraphs(cell)[0]!.elements[1]!.tooltip).toBeUndefined();
  setTableCellClickAction(cell, null, { range });
  expect(getTableCellParagraphs(cell)[0]!.elements.every((e) => !e.clickAction && !e.tooltip)).toBe(
    true,
  );
});

it('rejects invalid ranges and foreign slide targets without changing cell XML or relationships', async () => {
  const { pres, cell } = await fixture();
  const foreign = await fixture();
  const before = await savePresentation(pres);
  expect(() =>
    setTableCellClickAction(
      cell,
      { kind: 'url', url: 'https://example.com' },
      { range: { start: -1, end: 2 } },
    ),
  ).toThrow();
  expect(() => setTableCellClickAction(cell, { kind: 'slide', slide: foreign.slide })).toThrow();
  setTableCellClickAction(
    cell,
    { kind: 'url', url: 'https://example.com' },
    { range: { start: 2, end: 2 } },
  );
  expect(await savePresentation(pres)).toEqual(before);
});

it.each(['svg', 'foreignObject'] as const)(
  'renders cell URL and slide links only on the selected text (%s)',
  async (textLayout) => {
    const { pres, slide, cell } = await fixture();
    setTableCellClickAction(
      cell,
      { kind: 'url', url: 'https://example.com/?a=1&b=2' },
      { range: { start: 7, end: 10 }, tooltip: '参考 "Link" & <資料>' },
    );
    let svg = renderSlideToSvg(pres, slide, { textLayout });
    expect(svg).toContain('href="https://example.com/?a=1&amp;b=2"');
    expect(svg.match(/<a\s/g)).toHaveLength(1);
    if (textLayout === 'foreignObject')
      expect(svg).toContain('title="参考 &quot;Link&quot; &amp; &lt;資料&gt;"');
    setTableCellClickAction(
      cell,
      { kind: 'slide', slide: getSlides(pres)[1]! },
      { range: { start: 7, end: 10 } },
    );
    svg = renderSlideToSvg(pres, slide, { textLayout });
    expect(svg).toContain('href="#slide-2"');
    expect(svg.match(/<a\s/g)).toHaveLength(1);
  },
);
