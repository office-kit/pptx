// `findShapesInRect(slide, x, y, w, h)` — marquee-style region
// finder. Pair to the existing `findShapesAtPoint`.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  findShapesInRect,
  getShapeName,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findShapesInRect', () => {
  it('returns shapes overlapping the rectangle', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const a = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'A',
    });
    const b = addSlideTextBox(slide, {
      x: inches(3),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'B',
    });
    const c = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(3),
      w: inches(2),
      h: inches(1),
      text: 'C',
    });

    // Rectangle covers the top-left quadrant only — A is inside, B
    // straddles the right boundary, C is below.
    const found = findShapesInRect(slide, inches(0), inches(0), inches(3.5), inches(2));
    const names = found.map((s) => getShapeName(s));
    expect(names).toContain(getShapeName(a));
    expect(names).toContain(getShapeName(b));
    expect(names).not.toContain(getShapeName(c));
  });

  it('returns an empty array when the rectangle is empty', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'A',
    });
    // Rectangle far away from any shape.
    expect(findShapesInRect(slide, inches(100), inches(100), inches(1), inches(1))).toEqual([]);
  });

  it('counts touching edges as overlap', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const a = addSlideTextBox(slide, {
      x: inches(2),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      text: 'A',
    });
    // Rectangle just touches A's left edge at x=2".
    const found = findShapesInRect(slide, inches(0), inches(0), inches(2), inches(1));
    expect(found.map((s) => getShapeName(s))).toContain(getShapeName(a));
  });
});
