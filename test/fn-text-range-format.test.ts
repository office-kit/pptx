import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import {
  getSlides,
  getSlideShapes,
  getShapeText,
  getShapeParagraphElements,
  getShapeRunHyperlink,
  getSlideXmlString,
  loadPresentation,
  savePresentation,
  setShapeParagraphs,
  setShapeRunHyperlink,
  setShapeTextRangeFormat,
  replaceShapeTextRange,
} from '../src/api/index.ts';

async function fixture() {
  const presentation = await loadPresentation(
    await readFile(new URL('./fixtures/minimal/one-text-slide.pptx', import.meta.url)),
  );
  const slide = getSlides(presentation)[0]!;
  const shape = getSlideShapes(slide)[0]!;
  return { presentation, slide, shape };
}

describe('character-range formatting', () => {
  it('splits only selected text while preserving existing formats and hyperlinks after save', async () => {
    const { presentation, shape } = await fixture();
    setShapeParagraphs(shape, [
      {
        runs: [
          { text: 'Hello world', format: { italic: true, size: 24 } },
          { text: '!', format: { bold: true } },
        ],
      },
    ]);
    setShapeRunHyperlink(shape, 0, 0, 'https://example.com/');
    setShapeTextRangeFormat(shape, 6, 11, { bold: true, color: '#FF0000' });
    const loaded = await loadPresentation(await savePresentation(presentation));
    const result = getSlideShapes(getSlides(loaded)[0]!)[0]!;
    expect(getShapeText(result)).toBe('Hello world!');
    expect(getShapeParagraphElements(result, 0)).toMatchObject([
      { text: 'Hello ', format: { italic: true, size: 24 } },
      { text: 'world', format: { italic: true, size: 24, bold: true, color: '#FF0000' } },
      { text: '!', format: { bold: true } },
    ]);
    expect(getShapeParagraphElements(result, 0)[0]!.format?.bold).not.toBe(true);
    expect(getShapeRunHyperlink(result, 0, 0)).toBe('https://example.com/');
    expect(getShapeRunHyperlink(result, 0, 1)).toBe('https://example.com/');
    expect(getShapeRunHyperlink(result, 0, 2)).toBeNull();
  });

  it('uses UTF-16 offsets across paragraphs and preserves unselected Unicode', async () => {
    const { shape } = await fixture();
    setShapeParagraphs(shape, [{ runs: [{ text: 'A😀B' }] }, { runs: [{ text: '日本語' }] }]);
    setShapeTextRangeFormat(shape, 3, 7, { underline: 'sng' });
    expect(getShapeText(shape)).toBe('A😀B\n日本語');
    expect(getShapeParagraphElements(shape, 0)).toMatchObject([
      { text: 'A😀' },
      { text: 'B', format: { underline: true } },
    ]);
    expect(getShapeParagraphElements(shape, 1)).toMatchObject([
      { text: '日本', format: { underline: true } },
      { text: '語' },
    ]);
  });

  it('rejects invalid ranges atomically, including surrogate splits', async () => {
    const { slide, shape } = await fixture();
    setShapeParagraphs(shape, [{ runs: [{ text: 'A😀B' }] }]);
    const before = getSlideXmlString(slide);
    for (const [start, end] of [
      [-1, 1],
      [3, 2],
      [0, 99],
      [1, 2],
      [2, 3],
      [0.5, 1],
    ]) {
      expect(() => setShapeTextRangeFormat(shape, start!, end!, { bold: true })).toThrow();
      expect(getSlideXmlString(slide)).toBe(before);
    }
    setShapeTextRangeFormat(shape, 1, 1, { bold: true });
    expect(getSlideXmlString(slide)).toBe(before);
  });
});

describe('text range replacement', () => {
  it('preserves surrounding mixed formats and links through insert, delete and save', async () => {
    const { presentation, shape } = await fixture();
    setShapeParagraphs(shape, [
      {
        runs: [
          { text: 'Hello', format: { bold: true } },
          { text: ' world', format: { italic: true } },
        ],
      },
    ]);
    setShapeRunHyperlink(shape, 0, 1, 'https://example.com/');
    replaceShapeTextRange(shape, 2, 2, '😀');
    replaceShapeTextRange(shape, 4, 7, '');
    const loaded = await loadPresentation(await savePresentation(presentation));
    const result = getSlideShapes(getSlides(loaded)[0]!)[0]!;
    expect(getShapeText(result)).toBe('He😀 world');
    expect(getShapeParagraphElements(result, 0)).toMatchObject([
      { text: 'He', format: { bold: true } },
      { text: '😀', format: { bold: true } },
      { text: ' world', format: { italic: true } },
    ]);
    expect(getShapeRunHyperlink(result, 0, 2)).toBe('https://example.com/');
  });

  it('splits and joins paragraphs without replacing unaffected runs', async () => {
    const { shape } = await fixture();
    setShapeParagraphs(shape, [
      { align: 'ctr', runs: [{ text: 'AB', format: { bold: true } }] },
      { runs: [{ text: 'CD', format: { italic: true } }] },
      { runs: [{ text: 'Keep', format: { size: 42 } }] },
    ]);
    replaceShapeTextRange(shape, 1, 4, 'x\ny');
    expect(getShapeText(shape)).toBe('Ax\nyD\nKeep');
    expect(getShapeParagraphElements(shape, 1)).toMatchObject([
      { text: 'y', format: { bold: true } },
      { text: 'D', format: { italic: true } },
    ]);
    expect(getShapeParagraphElements(shape, 2)).toMatchObject([
      { text: 'Keep', format: { size: 42 } },
    ]);
    replaceShapeTextRange(shape, 2, 3, '');
    expect(getShapeText(shape)).toBe('AxyD\nKeep');
    replaceShapeTextRange(shape, 0, getShapeText(shape).length, '');
    expect(getShapeText(shape)).toBe('');
    replaceShapeTextRange(shape, 0, 0, 'New');
    expect(getShapeText(shape)).toBe('New');
  });

  it('uses the selected text style when replacing at a run boundary', async () => {
    const { shape } = await fixture();
    setShapeParagraphs(shape, [
      {
        runs: [
          { text: 'AB', format: { bold: true } },
          { text: 'CD', format: { italic: true } },
        ],
      },
    ]);
    replaceShapeTextRange(shape, 2, 4, 'X');
    expect(getShapeParagraphElements(shape, 0)).toMatchObject([
      { text: 'AB', format: { bold: true } },
      { text: 'X', format: { italic: true } },
    ]);
    expect(getShapeParagraphElements(shape, 0)[1]!.format?.bold).not.toBe(true);
  });

  it('rejects a surrogate split without changing the slide', async () => {
    const { slide, shape } = await fixture();
    setShapeParagraphs(shape, [{ runs: [{ text: 'A😀B' }] }]);
    const before = getSlideXmlString(slide);
    expect(() => replaceShapeTextRange(shape, 2, 3, 'x')).toThrow();
    expect(getSlideXmlString(slide)).toBe(before);
  });
});
