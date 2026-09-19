// `setShapeParagraphs` — structured paragraphs with per-run formats, the
// one path that yields several differently formatted runs in one paragraph.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';
import {
  addSlideTable,
  getParagraphAlignment,
  getParagraphEndFormat,
  getShapeParagraphCount,
  getShapeParagraphElements,
  getShapeText,
  getSlideShapes,
  getSlides,
  getTableCell,
  getTableCellParagraphs,
  inches,
  loadPresentation,
  mergeTableCells,
  readPackagePart,
  savePresentation,
  setShapeParagraphs,
  setShapeText,
  setTableCellParagraphs,
  setTableCellText,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const decoder = new TextDecoder();
const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

describe('fn API: setShapeParagraphs', () => {
  it('writes paragraphs with mixed-format runs and reads them back after a save', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const shape = getSlideShapes(getSlides(pres)[0]!)[0]!;
    setShapeParagraphs(shape, [
      {
        align: 'ctr',
        runs: [
          { text: 'Bold lead', format: { bold: true, size: 14, color: '#1F2937' } },
          { text: ' then plain' },
        ],
      },
      { runs: [{ text: '• second' }] },
      { runs: [] },
    ]);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const again = getSlideShapes(getSlides(reloaded)[0]!)[0]!;
    expect(getShapeParagraphCount(again)).toBe(3);
    expect(getShapeText(again)).toBe('Bold lead then plain\n• second\n');
    expect(getParagraphAlignment(again, 0)).toBe('ctr');
    expect(getParagraphAlignment(again, 1)).toBeNull();
    const elements = getShapeParagraphElements(again, 0);
    expect(elements).toHaveLength(2);
    expect(elements[0]).toMatchObject({
      kind: 'r',
      text: 'Bold lead',
      format: { bold: true, size: 14, color: '#1F2937' },
    });
    expect(elements[1]).toMatchObject({ kind: 'r', text: ' then plain' });
    expect(getShapeParagraphElements(again, 2)).toHaveLength(0);
  });

  it('keeps a CR LF inside a run verbatim across a save', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const shape = getSlideShapes(getSlides(pres)[0]!)[0]!;
    setShapeParagraphs(shape, [
      { runs: [{ text: 'Hello', format: { bold: true } }, { text: '\r\n' }, { text: 'world' }] },
    ]);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const again = getSlideShapes(getSlides(reloaded)[0]!)[0]!;
    expect(getShapeParagraphCount(again)).toBe(1);
    expect(
      getShapeParagraphElements(again, 0).map((e) => (e.kind === 'br' ? '<br>' : e.text)),
    ).toEqual(['Hello', '\r\n', 'world']);
  });
});

describe('fn API: setShapeParagraphs — empty input', () => {
  it('rejects an empty paragraph list', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const shape = getSlideShapes(getSlides(pres)[0]!)[0]!;
    expect(() => setShapeParagraphs(shape, [])).toThrow(/at least one paragraph/);
  });

  skipIfNoXmllint('accepts one empty paragraph as the empty body', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const shape = getSlideShapes(getSlides(pres)[0]!)[0]!;
    setShapeParagraphs(shape, [{ runs: [] }]);
    const reloaded = await loadPresentation(await savePresentation(pres));
    const again = getSlideShapes(getSlides(reloaded)[0]!)[0]!;
    expect(getShapeParagraphCount(again)).toBe(1);
    expect(getShapeText(again)).toBe('');
    expectSchemaValid(decoder.decode(readPackagePart(reloaded, '/ppt/slides/slide1.xml')!), 'pml');
  });
});

describe('fn API: setTableCellParagraphs', () => {
  it('writes several runs into one cell paragraph and reads them back after a save', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addSlideTable(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(1),
      rows: [['head'], ['body']],
    });
    setTableCellParagraphs(getTableCell(table, 1, 0), [
      {
        align: 'r',
        runs: [
          { text: 'alpha ', format: { bold: true, size: 9 } },
          { text: 'beta ', format: { bold: true, size: 9 } },
          { text: 'gamma', format: { bold: true, size: 9 } },
        ],
      },
    ]);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const cell = getTableCell(getSlideShapes(getSlides(reloaded)[0]!).at(-1)!, 1, 0);
    const paragraphs = getTableCellParagraphs(cell);
    expect(paragraphs).toHaveLength(1);
    expect(paragraphs[0]!.align).toBe('right');
    expect(paragraphs[0]!.elements.map((e) => (e.kind === 'br' ? '<br>' : e.text))).toEqual([
      'alpha ',
      'beta ',
      'gamma',
    ]);
    expect(paragraphs[0]!.elements[1]).toMatchObject({ format: { bold: true, size: 9 } });
  });
});

describe('fn API: setTableCellParagraphs — empty input', () => {
  const tableOnFreshDeck = async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const table = addSlideTable(getSlides(pres)[0]!, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(1),
      rows: [['head'], ['body']],
    });
    return { pres, cell: getTableCell(table, 1, 0) };
  };

  it('rejects an empty paragraph list', async () => {
    const { cell } = await tableOnFreshDeck();
    expect(() => setTableCellParagraphs(cell, [])).toThrow(/at least one paragraph/);
  });

  skipIfNoXmllint('accepts one empty paragraph as the empty cell', async () => {
    const { pres, cell } = await tableOnFreshDeck();
    setTableCellParagraphs(cell, [{ runs: [] }]);
    const reloaded = await loadPresentation(await savePresentation(pres));
    const table = getSlideShapes(getSlides(reloaded)[0]!).at(-1)!;
    expect(getTableCellParagraphs(getTableCell(table, 1, 0))).toEqual([
      { align: null, elements: [], endFormat: null },
    ]);
    expectSchemaValid(decoder.decode(readPackagePart(reloaded, '/ppt/slides/slide1.xml')!), 'pml');
  });
});

describe('fn API: paragraph end format (<a:endParaRPr>)', () => {
  const END_FORMAT = { size: 9, font: 'Yu Gothic', fontEastAsian: 'Yu Gothic' } as const;

  skipIfNoXmllint('writes endFormat after the runs and reads it back on a shape', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const shape = getSlideShapes(getSlides(pres)[0]!)[0]!;
    setShapeParagraphs(shape, [
      { align: 'ctr', runs: [{ text: 'Lead', format: { bold: true } }], endFormat: END_FORMAT },
      { runs: [], endFormat: { size: 24 } },
      { runs: [{ text: 'no end mark' }] },
    ]);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const again = getSlideShapes(getSlides(reloaded)[0]!)[0]!;
    expect(getParagraphEndFormat(again, 0)).toEqual(END_FORMAT);
    expect(getParagraphEndFormat(again, 1)).toEqual({ size: 24 });
    expect(getParagraphEndFormat(again, 2)).toBeNull();
    // The end mark is not an inline element: run readers must not see it.
    expect(getShapeParagraphElements(again, 1)).toHaveLength(0);

    const xml = decoder.decode(readPackagePart(reloaded, '/ppt/slides/slide1.xml')!);
    expect(xml).toContain(
      '</a:r><a:endParaRPr sz="900"><a:latin typeface="Yu Gothic"/><a:ea typeface="Yu Gothic"/></a:endParaRPr></a:p>',
    );
    expect(xml).toContain('<a:p><a:endParaRPr sz="2400"/></a:p>');
    expectSchemaValid(xml, 'pml');
  });

  skipIfNoXmllint('writes endFormat into an empty table cell and reads it back', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const table = addSlideTable(getSlides(pres)[0]!, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(1),
      rows: [['head'], ['body']],
    });
    setTableCellParagraphs(getTableCell(table, 0, 0), [
      { align: 'ctr', runs: [], endFormat: END_FORMAT },
    ]);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const again = getSlideShapes(getSlides(reloaded)[0]!).at(-1)!;
    expect(getTableCellParagraphs(getTableCell(again, 0, 0))).toEqual([
      { align: 'center', elements: [], endFormat: END_FORMAT },
    ]);
    const xml = decoder.decode(readPackagePart(reloaded, '/ppt/slides/slide1.xml')!);
    expect(xml).toContain('<a:p><a:pPr algn="ctr"/><a:endParaRPr sz="900">');
    expectSchemaValid(xml, 'pml');
  });

  it.each([Number.NaN, Number.POSITIVE_INFINITY, -1, 0, 4001])(
    'rejects end-mark size %s like a run size',
    async (size) => {
      const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
      const shape = getSlideShapes(getSlides(pres)[0]!)[0]!;
      expect(() => setShapeParagraphs(shape, [{ runs: [], endFormat: { size } }])).toThrow(
        RangeError,
      );
    },
  );

  skipIfNoXmllint('replaces an existing <a:endParaRPr> instead of adding a second', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const shape = getSlideShapes(getSlides(pres)[0]!)[0]!;
    setShapeParagraphs(shape, [{ runs: [{ text: 'x' }], endFormat: { size: 24 } }]);
    setShapeParagraphs(shape, [{ runs: [{ text: 'x' }], endFormat: { size: 12 } }]);

    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getParagraphEndFormat(getSlideShapes(getSlides(reloaded)[0]!)[0]!, 0)).toEqual({
      size: 12,
    });
    const xml = decoder.decode(readPackagePart(reloaded, '/ppt/slides/slide1.xml')!);
    expect(xml.match(/<a:endParaRPr/g)).toHaveLength(1);
    expectSchemaValid(xml, 'pml');
  });

  // A format rejected mid-list must not leave the body half replaced (or a
  // <a:txBody> with no <a:p>, which CT_TextBody forbids).
  skipIfNoXmllint.each([
    ['endFormat', { runs: [], endFormat: { size: Number.NaN } }],
    ['run format', { runs: [{ text: 'bad', format: { size: 4001 } }] }],
  ] as const)('a rejected %s leaves the existing text untouched', async (_name, invalid) => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    const table = addSlideTable(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(1),
      rows: [['head', 'covered', 'other']],
    });
    const original = [{ runs: [{ text: 'kept' }], endFormat: END_FORMAT }];
    setShapeParagraphs(shape, original);
    setTableCellParagraphs(getTableCell(table, 0, 0), original);
    mergeTableCells(table, { row: 0, col: 0, rowSpan: 1, colSpan: 2 }, { coveredText: 'drop' });
    const headBefore = getTableCellParagraphs(getTableCell(table, 0, 0));

    const rejected = [{ runs: [{ text: 'first is fine' }] }, invalid];
    expect(() => setShapeParagraphs(shape, rejected)).toThrow(RangeError);
    expect(() => setTableCellParagraphs(getTableCell(table, 0, 0), rejected)).toThrow(RangeError);
    expect(() => setTableCellParagraphs(getTableCell(table, 0, 1), rejected)).toThrow(RangeError);
    // The slide part is only re-serialized by a later successful edit.
    setTableCellText(getTableCell(table, 0, 2), 'edited');

    const reloaded = await loadPresentation(await savePresentation(pres));
    const againShape = getSlideShapes(getSlides(reloaded)[0]!)[0]!;
    const againTable = getSlideShapes(getSlides(reloaded)[0]!).at(-1)!;
    expect(getShapeText(againShape)).toBe('kept');
    expect(getParagraphEndFormat(againShape, 0)).toEqual(END_FORMAT);
    expect(getTableCellParagraphs(getTableCell(againTable, 0, 0))).toEqual(headBefore);
    expect(getTableCellParagraphs(getTableCell(againTable, 0, 1))).toEqual([]);
    expectSchemaValid(decoder.decode(readPackagePart(reloaded, '/ppt/slides/slide1.xml')!), 'pml');
  });

  it('setShapeText and setTableCellText do not keep the end-mark format', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    const cell = getTableCell(
      addSlideTable(slide, {
        x: inches(0.5),
        y: inches(0.5),
        w: inches(6),
        h: inches(1),
        rows: [['head']],
      }),
      0,
      0,
    );
    setShapeParagraphs(shape, [{ runs: [{ text: 'x' }], endFormat: END_FORMAT }]);
    setTableCellParagraphs(cell, [{ runs: [{ text: 'x' }], endFormat: END_FORMAT }]);

    setShapeText(shape, 'y');
    setTableCellText(cell, 'y');

    expect(getParagraphEndFormat(shape, 0)).toBeNull();
    expect(getTableCellParagraphs(cell)[0]!.endFormat).toBeNull();
  });

  it('throws on an out-of-range paragraph index', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const shape = getSlideShapes(getSlides(pres)[0]!)[0]!;
    expect(() => getParagraphEndFormat(shape, 99)).toThrow(RangeError);
  });
});
