// getSlideBoundingBox — union bounds across all positioned shapes.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideShape,
  findSlideLayout,
  getSlideBoundingBox,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideBoundingBox', () => {
  it('unions every positioned shape on the slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    addSlideShape(slide, {
      preset: 'rect', x: inches(1), y: inches(2), w: inches(3), h: inches(1),
    });
    addSlideShape(slide, {
      preset: 'ellipse', x: inches(5), y: inches(1), w: inches(2), h: inches(4),
    });
    const bb = getSlideBoundingBox(slide)!;
    expect(bb.x).toBe(inches(1));
    expect(bb.y).toBe(inches(1));
    // Right edge: max(1+3, 5+2) = 7
    expect(bb.w).toBe(inches(7 - 1));
    // Bottom edge: max(2+1, 1+4) = 5
    expect(bb.h).toBe(inches(5 - 1));
  });

  it('returns null when no shapes have bounds', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    // Blank layout has only dt/sldNum/ftr placeholders with no xfrm
    // visible on the slide itself, so the union is empty.
    expect(getSlideBoundingBox(slide)).toBeNull();
  });
});
