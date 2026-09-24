import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import * as pptx from '../src/api/index.ts';
import { applyBulletToParagraph } from '../src/internal/drawingml/text-body-mutation.ts';
import { NS, elem, qname } from '../src/internal/xml/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

it('edits a table cell range while retaining surrounding runs, paragraph formats and other cells', async () => {
  const p = await pptx.loadPresentation(
    await readFile(new URL('./fixtures/minimal/one-text-slide.pptx', import.meta.url)),
  );
  const slide = pptx.getSlides(p)[0]!;
  const table = pptx.addSlideTable(slide, {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(4),
    h: pptx.inches(2),
    rows: [
      ['', 'untouched'],
      ['lower', '😀'],
    ],
  });
  const cell = pptx.getTableCell(table, 0, 0);
  pptx.setTableCellParagraphs(cell, [
    {
      align: 'center',
      runs: [
        { text: 'Bold ', format: { bold: true } },
        { text: 'red', format: { color: 'FF0000' } },
      ],
      endFormat: { size: 22 },
    },
    { align: 'right', runs: [{ text: 'Tail', format: { italic: true } }] },
  ]);
  pptx.replaceTableCellTextRange(cell, 5, 8, 'new\nline');
  const check = (table: pptx.SlideShapeData) => {
    const changed = pptx.getTableCell(table, 0, 0);
    expect(pptx.getTableCellText(changed)).toBe('Bold new\nline\nTail');
    const paragraphs = pptx.getTableCellParagraphs(changed);
    expect(paragraphs.map((p) => p.align)).toEqual(['center', 'center', 'right']);
    expect(paragraphs[0]!.elements[0]!.format?.bold).toBe(true);
    expect(paragraphs[0]!.elements[1]!.format?.color).toBe('#FF0000');
    expect(paragraphs[1]!.elements[0]!.format?.color).toBe('#FF0000');
    expect(paragraphs[0]!.endFormat?.size).toBe(22);
    expect(paragraphs[2]!.elements[0]!.format?.italic).toBe(true);
    expect(pptx.getTableCellText(pptx.getTableCell(table, 0, 1))).toBe('untouched');
    expect(pptx.getTableCellText(pptx.getTableCell(table, 1, 0))).toBe('lower');
  };
  check(table);
  const loaded = await pptx.loadPresentation(await pptx.savePresentation(p));
  check(pptx.getSlideTables(pptx.getSlides(loaded)[0]!)[0]!);
  const emoji = pptx.getTableCell(table, 1, 1);
  const before = pptx.getSlideXmlString(slide);
  for (const [start, end] of [
    [-1, 0],
    [0, 3],
    [1, 1],
    [1.5, 2],
  ]) {
    expect(() => pptx.replaceTableCellTextRange(emoji, start!, end!, 'x')).toThrow();
    expect(pptx.getSlideXmlString(slide)).toBe(before);
  }
  pptx.replaceTableCellTextRange(emoji, 0, 2, '');
  expect(pptx.getTableCellText(emoji)).toBe('');
  pptx.replaceTableCellTextRange(emoji, 0, 0, '日本語');
  expect(pptx.getTableCellText(emoji)).toBe('日本語');
});

it('formats only the selected cell range and preserves its surrounding XML after reload', async () => {
  const p = await pptx.loadPresentation(
    await readFile(new URL('./fixtures/minimal/one-text-slide.pptx', import.meta.url)),
  );
  const slide = pptx.getSlides(p)[0]!;
  const table = pptx.addSlideTable(slide, {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(4),
    h: pptx.inches(2),
    rows: [['Hello 😀', 'other']],
  });
  const cell = pptx.getTableCell(table, 0, 0);
  pptx.setTableCellTextFormat(cell, { color: 'FF0000' });
  pptx.setTableCellTextRangeFormat(cell, 1, 4, { bold: true });
  const before = pptx.getSlideXmlString(slide);
  expect(() => pptx.setTableCellTextRangeFormat(cell, 7, 8, { italic: true })).toThrow();
  expect(() => pptx.setTableCellTextRangeFormat(cell, 0, 2, { size: -1 })).toThrow();
  expect(pptx.getSlideXmlString(slide)).toBe(before);
  const loaded = await pptx.loadPresentation(await pptx.savePresentation(p));
  const cells = pptx.getTableCells(pptx.getSlideTables(pptx.getSlides(loaded)[0]!)[0]!);
  const runs = pptx.getTableCellParagraphs(cells[0]![0]!)[0]!.elements;
  expect(runs.map((r) => (r.kind === 'br' ? '\n' : r.text))).toEqual(['H', 'ell', 'o 😀']);
  expect(runs.map((r) => !!r.format?.bold)).toEqual([false, true, false]);
  expect(runs.map((r) => r.format?.color)).toEqual(['#FF0000', '#FF0000', '#FF0000']);
  expect(pptx.getTableCellText(cells[0]![1]!)).toBe('other');
});

it('aligns selected cell paragraphs including carets, empty paragraphs and selection boundaries', async () => {
  const p = await pptx.loadPresentation(
    await readFile(new URL('./fixtures/minimal/one-text-slide.pptx', import.meta.url)),
  );
  const slide = pptx.getSlides(p)[0]!;
  const table = pptx.addSlideTable(slide, {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(4),
    h: pptx.inches(2),
    rows: [['', 'other']],
  });
  const cell = pptx.getTableCell(table, 0, 0);
  pptx.setTableCellParagraphs(cell, [
    { runs: [{ text: 'One', format: { bold: true } }, { text: '\n' }, { text: 'line' }] },
    { runs: [] },
    { runs: [{ text: 'Three 😀' }] },
  ]);
  expect(pptx.getTableCellText(cell)).toBe('One\nline\n\nThree 😀');
  const alignments = () => pptx.getTableCellParagraphs(cell).map((p) => p.align);
  pptx.setTableCellTextRangeAlignment(cell, 4, 4, 'center');
  expect(alignments()).toEqual(['center', null, null]);
  pptx.setTableCellTextRangeAlignment(cell, 9, 9, 'right');
  expect(alignments()).toEqual(['center', 'right', null]);
  pptx.setTableCellTextRangeAlignment(cell, 0, 10, 'justify');
  expect(alignments()).toEqual(['justify', 'justify', null]);
  pptx.setTableCellTextRangeAlignment(cell, 10, 10, 'distribute');
  expect(alignments()).toEqual(['justify', 'justify', 'distribute']);
  const before = pptx.getSlideXmlString(slide);
  expect(() => pptx.setTableCellTextRangeAlignment(cell, 17, 17, 'left')).toThrow();
  expect(() => pptx.setTableCellTextRangeAlignment(cell, -1, 0, 'left')).toThrow();
  expect(pptx.getSlideXmlString(slide)).toBe(before);
  const loaded = await pptx.loadPresentation(await pptx.savePresentation(p));
  const loadedTable = pptx.getSlideTables(pptx.getSlides(loaded)[0]!)[0]!;
  const paras = pptx.getTableCellParagraphs(pptx.getTableCell(loadedTable, 0, 0));
  expect(paras.map((p) => p.align)).toEqual(['justify', 'justify', 'distribute']);
  expect(paras[0]!.elements[0]!.format?.bold).toBe(true);
  expect(pptx.getTableCellText(pptx.getTableCell(loadedTable, 0, 1))).toBe('other');
});

it('changes cell line spacing at range boundaries, validates atomically and survives reload', async () => {
  const p = pptx.createPresentation();
  const slide = pptx.addBlankSlide(p);
  const table = pptx.addSlideTable(slide, {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(4),
    h: pptx.inches(2),
    rows: [['One\nTwo 😀\nThree', 'other']],
  });
  const cell = pptx.getTableCell(table, 0, 0);
  pptx.setTableCellTextRangeLineSpacing(cell, 4, 11, { kind: 'pct', value: 1.5 });
  const before = pptx.getSlideXmlString(slide);
  expect(() =>
    pptx.setTableCellTextRangeLineSpacing(cell, 9, 9, { kind: 'pts', value: 20 }),
  ).toThrow();
  expect(() =>
    pptx.setTableCellTextRangeLineSpacing(cell, 0, 0, { kind: 'pct', value: NaN }),
  ).toThrow();
  expect(pptx.getSlideXmlString(slide)).toBe(before);
  const loaded = await pptx.loadPresentation(await pptx.savePresentation(p));
  const restored = pptx.getTableCell(pptx.getSlideTables(pptx.getSlides(loaded)[0]!)[0]!, 0, 0);
  expect(
    pptx.getTableCellParagraphs(restored).map((p) => p.properties?.lineSpacing ?? null),
  ).toEqual([null, { kind: 'pct', value: 1.5 }, null]);
  pptx.setTableCellTextRangeLineSpacing(restored, 11, 11, { kind: 'pts', value: 22.5 });
  expect(pptx.getTableCellParagraphs(restored)[2]!.properties?.lineSpacing).toEqual({
    kind: 'pts',
    value: 22.5,
  });
  pptx.setTableCellTextRangeLineSpacing(restored, 4, 4, null);
  expect(pptx.getTableCellParagraphs(restored)[1]!.properties?.lineSpacing).toBeUndefined();
});

it('numbers selected cell paragraphs, preserves adjacent text and renders both preview modes', async () => {
  const p = pptx.createPresentation();
  const slide = pptx.addBlankSlide(p);
  const table = pptx.addSlideTable(slide, {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(6),
    h: pptx.inches(3),
    rows: [['One\nTwo 😀\nThree', 'other']],
  });
  const cell = pptx.getTableCell(table, 0, 0);
  pptx.setTableCellTextRangeBullets(cell, 0, 11, 'number');
  const before = pptx.getSlideXmlString(slide);
  expect(() => pptx.setTableCellTextRangeBullets(cell, 9, 9, 'bullet')).toThrow();
  expect(() => pptx.setTableCellTextRangeBullets(cell, 0, 0, { autoNum: 'invalid' })).toThrow();
  expect(pptx.getSlideXmlString(slide)).toBe(before);
  const restored = await pptx.loadPresentation(await pptx.savePresentation(p));
  const restoredSlide = pptx.getSlides(restored)[0]!;
  const restoredTable = pptx.getSlideTables(restoredSlide)[0]!;
  const restoredCell = pptx.getTableCell(restoredTable, 0, 0);
  expect(
    pptx.getTableCellParagraphs(restoredCell).map((p) => p.properties?.bullet ?? null),
  ).toEqual(['number', 'number', 'none']);
  expect(pptx.getTableCellText(pptx.getTableCell(restoredTable, 0, 1))).toBe('other');
  for (const textLayout of ['svg', 'foreignObject'] as const) {
    const svg = renderSlideToSvg(restored, restoredSlide, { textLayout });
    expect(svg).toContain('>1.<');
    expect(svg).toContain('>2.<');
    expect(svg).not.toContain('>3.<');
  }
  pptx.setTableCellTextRangeBullets(restoredCell, 4, 4, 'bullet');
  expect(pptx.getTableCellParagraphs(restoredCell)[1]!.properties?.bullet).toBe('bullet');
  pptx.setTableCellTextRangeBullets(restoredCell, 0, 11, 'none');
  expect(
    pptx.getTableCellParagraphs(restoredCell).map((p) => p.properties?.bullet ?? null),
  ).toEqual(['none', 'none', 'none']);
});

it('replaces image bullets and preserves DrawingML child order before default run properties', () => {
  const element = (name: string) => elem(qname('a', name, NS.dml));
  const pPr = element('pPr');
  pPr.children = ['lnSpc', 'buFontTx', 'buBlip', 'tabLst', 'defRPr', 'extLst'].map(element);
  const paragraph = element('p');
  paragraph.children = [pPr];
  applyBulletToParagraph(paragraph, 'number');
  const names = () =>
    pPr.children.map((child) => (child.kind === 'element' ? child.name.localName : ''));
  expect(names()).toEqual(['lnSpc', 'buFont', 'buAutoNum', 'tabLst', 'defRPr', 'extLst']);
  applyBulletToParagraph(paragraph, 'none');
  expect(names()).toEqual(['lnSpc', 'buNone', 'tabLst', 'defRPr', 'extLst']);
});

it('whole-cell formatting covers empty paragraphs and retains their existing formats', async () => {
  const presentation = pptx.createPresentation();
  const table = pptx.addSlideTable(pptx.addBlankSlide(presentation), {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(3),
    h: pptx.inches(1),
    rows: [['']],
  });
  const cell = pptx.getTableCell(table, 0, 0);
  pptx.setTableCellParagraphs(cell, [
    { runs: [], endFormat: { italic: true } },
    { runs: [{ text: 'Keep', format: { underline: true } }] },
  ]);
  pptx.setTableCellTextFormat(cell, { bold: true, size: 24 });
  const saved = await pptx.loadPresentation(await pptx.savePresentation(presentation));
  const paragraphs = pptx.getTableCellParagraphs(
    pptx.getTableCell(pptx.getSlideTables(pptx.getSlides(saved)[0]!)[0]!, 0, 0),
  );
  expect(paragraphs[0]!.endFormat).toMatchObject({ italic: true, bold: true, size: 24 });
  expect(paragraphs[1]!.endFormat).toMatchObject({ bold: true, size: 24 });
  expect(paragraphs[1]!.elements[0]!.format).toMatchObject({
    bold: true,
    size: 24,
    underline: true,
  });
});
