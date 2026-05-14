// Level-3 free-function deck manipulation.
//
// Verifies that the tree-shakeable `fn.ts` API produces the same
// observable behavior as the class-based API for the deck-level
// operations: add, remove, move, duplicate, getSlides, getSlideText,
// getSlideLayouts, replaceTokensInPresentation.
//
// The class API is still used to validate the *result* of free-function
// edits, because slide-level mutation (setText, etc.) hasn't migrated
// yet — both APIs share opaque internal state, so a slide produced by
// `addSlide(...)` can be re-read by the class API to confirm placement,
// layout binding, and slide count.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  Presentation,
  addSlide,
  duplicateSlide,
  findSlideLayout,
  getSlideLayouts,
  getSlideText,
  getSlides,
  loadPresentation,
  moveSlide,
  removeSlide,
  replaceTokensInPresentation,
  savePresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: deck manipulation', () => {
  it('getSlides + getSlideText read existing slides', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slides = getSlides(pres);
    expect(slides.length).toBe(2);
    expect(getSlideText(slides[0]!).length).toBeGreaterThan(0);
  });

  it('getSlideLayouts enumerates layouts in the package', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layouts = getSlideLayouts(pres);
    expect(layouts.length).toBeGreaterThan(0);
  });

  it('addSlide appends a slide bound to the chosen layout', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    expect(getSlides(pres).length).toBe(0);

    const layout = findSlideLayout(pres, 'Title and Content');
    expect(layout).not.toBeNull();
    const slide = addSlide(pres, { layout: layout! });
    expect(slide).toBeDefined();
    expect(getSlides(pres).length).toBe(1);

    // Round-trip and verify the new slide's layout binding via the class API.
    const reloaded = await Presentation.load(await savePresentation(pres));
    expect(reloaded.slides.length).toBe(1);
    expect(reloaded.slides[0]?.layout?.name).toBe('Title and Content');
  });

  it('removeSlide drops the slide and updates rels + sldIdLst', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slides = getSlides(pres);
    const target = slides[0]!;
    removeSlide(pres, target);

    expect(getSlides(pres).length).toBe(1);
    const reloaded = await Presentation.load(await savePresentation(pres));
    expect(reloaded.slides.length).toBe(1);
  });

  it('moveSlide reorders by sldIdLst index without changing identity', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const before = getSlides(pres).map((s) => getSlideText(s));
    expect(before.length).toBe(2);

    const first = getSlides(pres)[0]!;
    moveSlide(pres, first, 1);
    const after = getSlides(pres).map((s) => getSlideText(s));
    expect(after).toEqual([before[1], before[0]]);

    const reloaded = await Presentation.load(await savePresentation(pres));
    expect(reloaded.slides.map((s) => s.text)).toEqual([before[1], before[0]]);
  });

  it('duplicateSlide appends a deep copy of the source slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const source = getSlides(pres)[0]!;
    const sourceText = getSlideText(source);
    expect(sourceText.length).toBeGreaterThan(0);

    const dup = duplicateSlide(pres, source);
    expect(dup).toBeDefined();
    expect(getSlides(pres).length).toBe(2);
    expect(getSlideText(dup)).toBe(sourceText);
  });

  it('replaceTokensInPresentation replaces across slides', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    // Inject a token by replacing the existing text via the class API
    // (slide-level mutation hasn't migrated to fn yet).
    const classApi = await Presentation.load(await savePresentation(pres));
    const slide = classApi.slides[0];
    if (!slide) throw new Error('expected one slide');
    const shape = slide.shapes.find((s) => s.text.length > 0);
    if (!shape) throw new Error('expected one text shape');
    shape.setText('Hello, {{name}}!');

    const seeded = await loadPresentation(await classApi.save());
    const count = replaceTokensInPresentation(seeded, { name: 'World' });
    expect(count).toBeGreaterThan(0);
    expect(getSlideText(getSlides(seeded)[0]!)).toContain('Hello, World!');
  });
});
