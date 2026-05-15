// importSlide — copy a slide from one deck into another.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  Presentation,
  addSlide,
  addSlideImage,
  findSlideLayout,
  getMediaParts,
  getSlideShapes,
  getSlideText,
  getSlides,
  importSlide,
  inches,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const tinyPng = (): Uint8Array =>
  new Uint8Array([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x62, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
    0x42, 0x60, 0x82,
  ]);

describe('fn API: importSlide', () => {
  it('copies a slide from one deck to another', async () => {
    // Source deck.
    const source = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const sourceText = getSlideText(getSlides(source)[0]!);

    // Target deck.
    const target = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(target, 'Title and Content');
    expect(layout).not.toBeNull();

    const imported = importSlide(target, getSlides(source)[0]!, layout!);
    expect(imported).toBeDefined();
    expect(getSlideText(imported)).toBe(sourceText);
    expect(getSlides(target)).toHaveLength(1);

    // Round-trip — the target deck should still load cleanly.
    const reloaded = await Presentation.load(await savePresentation(target));
    expect(reloaded.slides.length).toBe(1);
  });

  it('copies image media along with the slide', async () => {
    const source = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    addSlideImage(getSlides(source)[0]!, tinyPng(), {
      x: inches(0), y: inches(0), w: inches(2), h: inches(2),
      format: 'png',
    });

    const target = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(target, 'Title and Content')!;
    expect(getMediaParts(target)).toEqual([]);

    importSlide(target, getSlides(source)[0]!, layout);

    const media = getMediaParts(target);
    expect(media.length).toBeGreaterThan(0);
    expect(media.some((m) => m.contentType.includes('png'))).toBe(true);
  });

  it('binds the imported slide to the supplied layout', async () => {
    const source = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const target = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(target, 'Blank')!;
    addSlide(target, { layout: findSlideLayout(target, 'Title Slide')! });

    importSlide(target, getSlides(source)[0]!, layout);

    const reloaded = await Presentation.load(await savePresentation(target));
    expect(reloaded.slides.at(-1)?.layout?.name).toBe('Blank');
    expect(getSlideShapes(getSlides(target).at(-1)!).length).toBeGreaterThan(0);
  });
});
