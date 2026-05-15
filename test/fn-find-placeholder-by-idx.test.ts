// findSlidePlaceholderByIdx — disambiguate same-type placeholders by
// their <p:ph idx="..."> attribute.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayout,
  findSlidePlaceholderByIdx,
  findSlidePlaceholders,
  getShapePlaceholderIdx,
  getShapeText,
  loadPresentation,
  setShapeText,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findSlidePlaceholderByIdx', () => {
  it('finds each body slot on a Two Content layout by idx', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Two Content')!;
    const slide = addSlide(pres, { layout });

    const bodies = findSlidePlaceholders(slide, 'body');
    // Sanity: two body slots, each with a distinct idx.
    expect(bodies.length).toBe(2);
    const indices = bodies.map((b) => getShapePlaceholderIdx(b)).sort();
    expect(indices.length).toBe(2);
    expect(indices[0]).not.toBe(indices[1]);

    // Disambiguate.
    const left = findSlidePlaceholderByIdx(slide, indices[0]!)!;
    const right = findSlidePlaceholderByIdx(slide, indices[1]!)!;
    setShapeText(left, 'L');
    setShapeText(right, 'R');
    expect(getShapeText(left)).toBe('L');
    expect(getShapeText(right)).toBe('R');
  });

  it('returns null for an idx that no shape carries', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout });
    expect(findSlidePlaceholderByIdx(slide, 99)).toBeNull();
  });
});
