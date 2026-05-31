// `isShapeTextBox` — distinguishes a text box (`<p:cNvSpPr txBox="1">`) from an
// autoshape. The distinction drives default text formatting (text boxes
// left/top, autoshapes center/middle), so it must be reliable.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  addSlideTextBox,
  addTitleSlide,
  getShapePlaceholderType,
  getSlideShapes,
  getSlides,
  inches,
  isShapeTextBox,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: isShapeTextBox', () => {
  it('is true for a text box and false for an autoshape', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'box',
    });
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(2),
      y: inches(0),
      w: inches(2),
      h: inches(1),
    });
    expect(isShapeTextBox(tb)).toBe(true);
    expect(isShapeTextBox(shape)).toBe(false);
  });

  it('is false for a placeholder', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const slide = addTitleSlide(pres, 'Hi');
    const ctr = getSlideShapes(slide).find((s) => getShapePlaceholderType(s) === 'ctrTitle');
    expect(ctr).toBeDefined();
    expect(isShapeTextBox(ctr!)).toBe(false);
  });
});
