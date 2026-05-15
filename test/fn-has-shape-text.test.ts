// hasShapeText — predicate for "this shape has visible text".

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  addSlideTextBox,
  clearShapeText,
  getSlides,
  hasShapeText,
  inches,
  loadPresentation,
  setShapeText,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: hasShapeText', () => {
  it('is true when a textbox has text', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(2), h: inches(1), text: 'hi',
    });
    expect(hasShapeText(tb)).toBe(true);
  });

  it('is false after clearing the text', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(2), h: inches(1), text: 'set',
    });
    expect(hasShapeText(tb)).toBe(true);
    clearShapeText(tb);
    expect(hasShapeText(tb)).toBe(false);
  });

  it('is false on a freshly-created plain shape', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const rect = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    expect(hasShapeText(rect)).toBe(false);
  });

  it('reflects writes through setShapeText on a textbox', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(2), h: inches(1), text: '',
    });
    expect(hasShapeText(tb)).toBe(false);
    setShapeText(tb, 'now I have text');
    expect(hasShapeText(tb)).toBe(true);
  });
});
