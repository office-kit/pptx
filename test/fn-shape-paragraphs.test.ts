// `setShapeParagraphs` — structured paragraphs with per-run formats, the
// one path that yields several differently formatted runs in one paragraph.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';
import {
  addSlideTable,
  getParagraphAlignment,
  getShapeParagraphCount,
  getShapeParagraphElements,
  getShapeText,
  getSlideShapes,
  getSlides,
  getTableCell,
  getTableCellParagraphs,
  inches,
  loadPresentation,
  readPackagePart,
  savePresentation,
  setShapeParagraphs,
  setTableCellParagraphs,
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
      { align: null, elements: [] },
    ]);
    expectSchemaValid(decoder.decode(readPackagePart(reloaded, '/ppt/slides/slide1.xml')!), 'pml');
  });
});
