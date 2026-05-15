// appendShapeText — add a paragraph to a shape without replacing
// existing text.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  appendShapeText,
  getShapeText,
  getSlides,
  inches,
  loadPresentation,
  setShapeText,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: appendShapeText', () => {
  it('writes the first line when the shape has no text yet', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(3), h: inches(2),
      text: '',
    });
    appendShapeText(tb, 'first');
    expect(getShapeText(tb)).toBe('first');
  });

  it('preserves existing text and adds a newline', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(3), h: inches(2),
      text: 'opening',
    });
    appendShapeText(tb, 'reminder');
    expect(getShapeText(tb)).toBe('opening\nreminder');

    appendShapeText(tb, 'last');
    expect(getShapeText(tb)).toBe('opening\nreminder\nlast');
  });

  it('chains after setShapeText', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(3), h: inches(2),
      text: '',
    });
    setShapeText(tb, 'A');
    appendShapeText(tb, 'B');
    setShapeText(tb, 'C');
    appendShapeText(tb, 'D');
    expect(getShapeText(tb)).toBe('C\nD');
  });

  it('throws on non-text shapes', async () => {
    const { addSlideImage } = await import('../src/api/index.ts');
    const PNG = new Uint8Array([
      0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
      0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
      0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
      0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
      0x42, 0x60, 0x82,
    ]);
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const pic = addSlideImage(slide, PNG, {
      x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    expect(() => appendShapeText(pic, 'x')).toThrow();
  });
});
