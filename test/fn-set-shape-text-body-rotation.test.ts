// `setShapeTextBodyRotationDeg` — text-body rotation via `<a:bodyPr rot/>`.
// Companion writer to `getShapeTextBodyRotationDeg`; rotates only the
// text inside the shape, not the shape itself.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  getShapeTextBodyRotationDeg,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeTextBodyRotationDeg,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: setShapeTextBodyRotationDeg', () => {
  it('round-trips a positive degree value through save/reload', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'rotated',
    });
    setShapeTextBodyRotationDeg(tb, 90);
    expect(getShapeTextBodyRotationDeg(tb)).toBe(90);

    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const reShapes = getSlideShapes(getSlides(reloaded)[0]!);
    const rTb = reShapes[reShapes.length - 1]!;
    expect(getShapeTextBodyRotationDeg(rTb)).toBe(90);
  });

  it('round-trips a negative degree value', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'r',
    });
    setShapeTextBodyRotationDeg(tb, -45);
    expect(getShapeTextBodyRotationDeg(tb)).toBe(-45);
  });

  it('clears the attribute when set to null or 0', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'r',
    });
    setShapeTextBodyRotationDeg(tb, 180);
    expect(getShapeTextBodyRotationDeg(tb)).toBe(180);
    setShapeTextBodyRotationDeg(tb, null);
    expect(getShapeTextBodyRotationDeg(tb)).toBeNull();
    setShapeTextBodyRotationDeg(tb, 90);
    setShapeTextBodyRotationDeg(tb, 0);
    expect(getShapeTextBodyRotationDeg(tb)).toBeNull();
  });
});
