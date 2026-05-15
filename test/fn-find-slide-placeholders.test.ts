// findSlidePlaceholders — multi-match variant of findSlidePlaceholder.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayout,
  findSlidePlaceholder,
  findSlidePlaceholders,
  getShapeText,
  loadPresentation,
  setShapeText,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findSlidePlaceholders', () => {
  it('finds all body placeholders on a Two-Content layout', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Two Content')!;
    const slide = addSlide(pres, { layout });
    const bodies = findSlidePlaceholders(slide, 'body');
    // PowerPoint's stock "Two Content" layout has two body slots.
    expect(bodies.length).toBe(2);
    // Set distinct text to confirm they're independent.
    setShapeText(bodies[0]!, 'left');
    setShapeText(bodies[1]!, 'right');
    expect(getShapeText(bodies[0]!)).toBe('left');
    expect(getShapeText(bodies[1]!)).toBe('right');
  });

  it('returns an empty array when no placeholder matches', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout });
    expect(findSlidePlaceholders(slide, 'title')).toEqual([]);
    expect(findSlidePlaceholders(slide, 'body')).toEqual([]);
  });

  it('agrees with findSlidePlaceholder on the first match', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    const slide = addSlide(pres, { layout });
    const all = findSlidePlaceholders(slide, 'title');
    const first = findSlidePlaceholder(slide, 'title');
    // Either both null or first === all[0].
    if (first === null) {
      expect(all).toEqual([]);
    } else {
      expect(all.length).toBeGreaterThan(0);
      expect(all[0]).toBe(first);
    }
  });
});
