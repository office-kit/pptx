// getShapeFillColor — sugar over getShapeFill for the solid case.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  getShapeFillColor,
  getSlides,
  inches,
  loadPresentation,
  setShapeFill,
  setShapeNoFill,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapeFillColor', () => {
  it('returns the solid color set via setShapeFill', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const rect = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    setShapeFill(rect, '#AABBCC');
    expect(getShapeFillColor(rect)).toBe('#AABBCC');
  });

  it('returns null for non-solid fills', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const rect = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    setShapeNoFill(rect);
    expect(getShapeFillColor(rect)).toBeNull();
  });
});
