// Free-function slide title convenience: getSlideTitle / setSlideTitle.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayout,
  findSlidePlaceholder,
  getShapeText,
  getSlideTitle,
  getSlides,
  loadPresentation,
  savePresentation,
  setSlideTitle,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: slide title convenience', () => {
  it('setSlideTitle + getSlideTitle round-trip on a Title and Content slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content');
    expect(layout).not.toBeNull();
    const slide = addSlide(pres, { layout: layout! });

    setSlideTitle(slide, 'Brand new slide');
    expect(getSlideTitle(slide)).toBe('Brand new slide');

    const reloaded = await loadPresentation(await savePresentation(pres));
    const title = findSlidePlaceholder(getSlides(reloaded)[0]!, 'title');
    expect(title && getShapeText(title)).toBe('Brand new slide');
  });

  it('getSlideTitle returns null when no title placeholder exists', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank');
    expect(layout).not.toBeNull();
    const slide = addSlide(pres, { layout: layout! });
    expect(getSlideTitle(slide)).toBeNull();
  });

  it('setSlideTitle throws on a slide with no title placeholder', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank');
    const slide = addSlide(pres, { layout: layout! });
    expect(() => setSlideTitle(slide, 'x')).toThrow(/no title/);
  });

  it('falls back to ctrTitle when the layout has no plain title', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title Slide');
    expect(layout).not.toBeNull();
    const slide = addSlide(pres, { layout: layout! });

    setSlideTitle(slide, 'Hero title');
    expect(getSlideTitle(slide)).toBe('Hero title');
  });

  it('findSlideLayout returns null for a missing layout name', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    expect(findSlideLayout(pres, 'No Such Layout')).toBeNull();
  });
});
