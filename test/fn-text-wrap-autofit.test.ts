// Text wrap + auto-fit on bodyPr.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  getShapeTextAutoFit,
  getShapeTextWrap,
  getSlides,
  inches,
  loadPresentation,
  setShapeTextAutoFit,
  setShapeTextWrap,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: setShapeTextWrap', () => {
  it('round-trips both wrap modes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(3), h: inches(2),
      text: 'A',
    });
    setShapeTextWrap(tb, 'none');
    expect(getShapeTextWrap(tb)).toBe('none');
    setShapeTextWrap(tb, 'square');
    expect(getShapeTextWrap(tb)).toBe('square');
  });
});

describe('fn API: setShapeTextAutoFit', () => {
  it('round-trips every mode', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(3), h: inches(2),
      text: 'A',
    });
    for (const mode of ['none', 'normal', 'shape'] as const) {
      setShapeTextAutoFit(tb, mode);
      expect(getShapeTextAutoFit(tb)).toBe(mode);
    }
  });

  it('replaces the prior auto-fit child each call', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(3), h: inches(2),
      text: 'A',
    });
    setShapeTextAutoFit(tb, 'normal');
    setShapeTextAutoFit(tb, 'shape');
    expect(getShapeTextAutoFit(tb)).toBe('shape');
  });
});
