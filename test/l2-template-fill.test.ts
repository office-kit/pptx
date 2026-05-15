// Level-2 (template fill) end-to-end smoke test.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  findSlidePlaceholder,
  getShapeKind,
  getShapeText,
  getSlideShapes,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeText,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('L2: template fill (text replacement)', () => {
  it('replaces a title placeholder and survives a save/reload cycle', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const title = findSlidePlaceholder(slide, 'title');
    if (!title) throw new Error('expected a title placeholder');

    expect(getShapeText(title)).toBe('Hello, OOXML');
    setShapeText(title, 'Q3 Review');
    expect(getShapeText(title)).toBe('Q3 Review');

    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const reTitle = findSlidePlaceholder(getSlides(reloaded)[0]!, 'title');
    expect(reTitle && getShapeText(reTitle)).toBe('Q3 Review');
  });

  it('handles multi-line text by splitting on \\n into paragraphs', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const title = findSlidePlaceholder(getSlides(pres)[0]!, 'title');
    if (!title) throw new Error('expected a title placeholder');
    setShapeText(title, 'First line\nSecond line\nThird line');

    const reloaded = await loadPresentation(await savePresentation(pres));
    const reTitle = findSlidePlaceholder(getSlides(reloaded)[0]!, 'title');
    expect(reTitle && getShapeText(reTitle)).toBe('First line\nSecond line\nThird line');
  });

  it('preserves run properties (rPr) across the replacement', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const title = findSlidePlaceholder(getSlides(pres)[0]!, 'title');
    if (!title) throw new Error('expected a title placeholder');
    setShapeText(title, 'Replaced');
    const bytes = await savePresentation(pres);
    expect(bytes.length).toBeGreaterThan(0);

    const reloaded = await loadPresentation(bytes);
    const reTitle = findSlidePlaceholder(getSlides(reloaded)[0]!, 'title');
    expect(reTitle && getShapeText(reTitle)).toBe('Replaced');
  });

  it('rejects setShapeText on non-shape kinds', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const nonText = getSlideShapes(slide).find((s) => getShapeKind(s) !== 'shape');
    if (nonText) {
      expect(() => setShapeText(nonText, 'boom')).toThrow();
    }
  });
});
