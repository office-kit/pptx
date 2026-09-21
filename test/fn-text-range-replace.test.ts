import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import {
  loadPresentation,
  getSlides,
  getSlideShapes,
  getShapeText,
  getShapeParagraphElements,
  setShapeParagraphs,
  setShapeText,
  savePresentation,
} from '../src/api/index.ts';

async function seeded() {
  const pres = await loadPresentation(
    await readFile(new URL('./fixtures/minimal/one-text-slide.pptx', import.meta.url)),
  );
  const shape = getSlideShapes(getSlides(pres)[0]!)[0]!;
  setShapeParagraphs(shape, [
    {
      runs: [
        { text: 'A', format: { bold: true } },
        { text: 'A', format: { italic: true } },
        { text: '😀後' },
      ],
    },
  ]);
  return { pres, shape };
}

describe('exact text replacement ranges', () => {
  it('deletes the selected identical character and retains the other character’s format', async () => {
    const { pres, shape } = await seeded();
    setShapeText(shape, '', { range: { start: 0, end: 1 } });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const result = getSlideShapes(getSlides(reloaded)[0]!)[0]!;
    expect(getShapeText(result)).toBe('A😀後');
    expect(getShapeParagraphElements(result, 0)[0]).toMatchObject({
      text: 'A',
      format: { italic: true },
    });
  });
  it('replaces exactly one selection with literal text', async () => {
    const { shape } = await seeded();
    setShapeText(shape, '$&日本語', { range: { start: 1, end: 4 } });
    expect(getShapeText(shape)).toBe('A$&日本語後');
    expect(getShapeParagraphElements(shape, 0)[0]).toMatchObject({
      text: 'A',
      format: { bold: true },
    });
    expect(getShapeParagraphElements(shape, 0)[1]).toMatchObject({
      text: '$&日本語',
      format: { italic: true },
    });
  });
  it.each([
    { start: -1, end: 1 },
    { start: 3, end: 4 },
    { start: 2, end: 3 },
    { start: 0, end: 99 },
    { start: 2, end: 1 },
    { start: 0.5, end: 1 },
  ])('rejects invalid ranges without changing text: %j', async (range) => {
    const { shape } = await seeded();
    expect(() => setShapeText(shape, 'bad', { range })).toThrow(RangeError);
    expect(getShapeText(shape)).toBe('AA😀後');
  });
});
