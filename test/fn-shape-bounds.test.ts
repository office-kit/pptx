// `getShapeBounds` / `setShapeBounds` — combined position + size accessor.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  getShapeBounds,
  getSlides,
  inches,
  loadPresentation,
  setShapeBounds,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapeBounds / setShapeBounds', () => {
  it('round-trips a fresh bounds rect', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0.5),
      y: inches(0.5),
      w: inches(3),
      h: inches(2),
    });
    const b = getShapeBounds(shape);
    expect(b).toEqual({ x: inches(0.5), y: inches(0.5), w: inches(3), h: inches(2) });
  });

  it('setShapeBounds writes both position and size in one call', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    setShapeBounds(shape, { x: inches(2), y: inches(3), w: inches(4), h: inches(5) });
    expect(getShapeBounds(shape)).toEqual({
      x: inches(2),
      y: inches(3),
      w: inches(4),
      h: inches(5),
    });
  });
});
