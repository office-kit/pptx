// Free-function slide title convenience: getSlideTitle / setSlideTitle.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  Presentation,
  addSlide,
  getSlideLayouts,
  getSlideTitle,
  loadPresentation,
  savePresentation,
  setSlideTitle,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: slide title convenience', () => {
  it('setSlideTitle + getSlideTitle round-trip on a Title and Content slide', async () => {
    // blank.pptx has no slides; add one with a `Title and Content` layout
    // (which has a `title` placeholder) then set + read the title.
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const reloadedForLayoutLookup = await Presentation.load(await savePresentation(pres));
    const targetIdx = reloadedForLayoutLookup.slideLayouts.findIndex(
      (l) => l.name === 'Title and Content',
    );
    expect(targetIdx).toBeGreaterThanOrEqual(0);
    const layouts = getSlideLayouts(pres);
    const layout = layouts[targetIdx]!;
    const slide = addSlide(pres, { layout });

    setSlideTitle(slide, 'Brand new slide');
    expect(getSlideTitle(slide)).toBe('Brand new slide');

    // Persists across save → reload.
    const reloaded = await Presentation.load(await savePresentation(pres));
    expect(reloaded.slides[0]?.findPlaceholder('title')?.text).toBe('Brand new slide');
  });

  it('getSlideTitle returns null when no title placeholder exists', async () => {
    // blank.pptx itself has no slides, so we use a fresh "Blank" layout
    // for a no-title slide.
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const reloadedForLayoutLookup = await Presentation.load(await savePresentation(pres));
    const targetIdx = reloadedForLayoutLookup.slideLayouts.findIndex(
      (l) => l.name === 'Blank',
    );
    expect(targetIdx).toBeGreaterThanOrEqual(0);
    const layouts = getSlideLayouts(pres);
    const slide = addSlide(pres, { layout: layouts[targetIdx]! });
    expect(getSlideTitle(slide)).toBeNull();
  });

  it('setSlideTitle throws on a slide with no title placeholder', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const reloadedForLayoutLookup = await Presentation.load(await savePresentation(pres));
    const targetIdx = reloadedForLayoutLookup.slideLayouts.findIndex(
      (l) => l.name === 'Blank',
    );
    const layouts = getSlideLayouts(pres);
    const slide = addSlide(pres, { layout: layouts[targetIdx]! });
    expect(() => setSlideTitle(slide, 'x')).toThrow(/no title/);
  });

  it('falls back to ctrTitle when the layout has no plain title', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const reloadedForLayoutLookup = await Presentation.load(await savePresentation(pres));
    const targetIdx = reloadedForLayoutLookup.slideLayouts.findIndex(
      (l) => l.name === 'Title Slide',
    );
    expect(targetIdx).toBeGreaterThanOrEqual(0);
    const layouts = getSlideLayouts(pres);
    const slide = addSlide(pres, { layout: layouts[targetIdx]! });

    setSlideTitle(slide, 'Hero title');
    expect(getSlideTitle(slide)).toBe('Hero title');
  });
});
