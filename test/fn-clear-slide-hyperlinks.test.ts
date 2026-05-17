// clearSlideHyperlinks — slide-level URL sanitizer.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import type { SlideData } from '../src/api/index.ts';
import {
  addSlide,
  addSlideTextBox,
  clearSlideHyperlinks,
  findSlideLayout,
  getShapeHyperlink,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  setShapeHyperlink,
} from '../src/api/index.ts';

const slideUrls = (slide: SlideData): string[] => {
  const out: string[] = [];
  for (const s of getSlideShapes(slide)) {
    const u = getShapeHyperlink(s);
    if (u !== null) out.push(u);
  }
  return out;
};

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: clearSlideHyperlinks', () => {
  it('clears only the targeted slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    {
      const slide = addSlide(pres, { layout: blank });
      const a = addSlideTextBox(slide, {
        x: inches(0),
        y: inches(0),
        w: inches(2),
        h: inches(1),
        text: 'a',
      });
      setShapeHyperlink(a, 'https://example.com/a');
    }
    const secondLinkUrl = 'https://example.com/b';
    {
      const slide = addSlide(pres, { layout: blank });
      const b = addSlideTextBox(slide, {
        x: inches(0),
        y: inches(0),
        w: inches(2),
        h: inches(1),
        text: 'b',
      });
      setShapeHyperlink(b, secondLinkUrl);
    }
    // Re-fetch slide handles — earlier ones may have been invalidated
    // by the second addSlide.
    const slides = getSlides(pres);
    expect(clearSlideHyperlinks(slides[0]!)).toBe(1);
    expect(slideUrls(slides[0]!)).toEqual([]);
    expect(slideUrls(slides[1]!)).toEqual([secondLinkUrl]);
  });

  it('returns 0 when slide has nothing to clear', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'plain',
    });
    expect(clearSlideHyperlinks(slide)).toBe(0);
  });
});
