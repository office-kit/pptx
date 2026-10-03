// isShapeHidden / setShapeHidden — toggle <p:cNvPr hidden="1">.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  getSlideShapes,
  getSlides,
  inches,
  isShapeHidden,
  loadPresentation,
  savePresentation,
  setShapeHidden,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: isShapeHidden / setShapeHidden', () => {
  it('a freshly-added shape is not hidden', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const s = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    expect(isShapeHidden(s)).toBe(false);
  });

  it('round-trips the hidden flag', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const s = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    setShapeHidden(s, true);
    expect(isShapeHidden(s)).toBe(true);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const { getSlideShapes } = await import('../src/api/index.ts');
    const shapes = getSlideShapes(getSlides(reloaded)[0]!);
    const last = shapes[shapes.length - 1]!;
    expect(isShapeHidden(last)).toBe(true);
  });

  it('toggles back to visible', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const s = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    setShapeHidden(s, true);
    expect(isShapeHidden(s)).toBe(true);
    setShapeHidden(s, false);
    expect(isShapeHidden(s)).toBe(false);
  });

  it('hides a batch across slides with one refresh per owning slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slides = getSlides(pres);
    const first = addSlideShape(slides[0]!, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    const second = addSlideShape(slides[1]!, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    setShapeHidden([first, second], true);
    expect(isShapeHidden(first)).toBe(true);
    expect(isShapeHidden(second)).toBe(true);
    expect(getSlideShapes(slides[0]!).some((shape) => isShapeHidden(shape))).toBe(true);
    expect(getSlideShapes(slides[1]!).some((shape) => isShapeHidden(shape))).toBe(true);
  });
});
