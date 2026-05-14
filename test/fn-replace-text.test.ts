// Free-text replace across the deck (no {{token}} braces required).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  Presentation,
  getSlideShapes,
  getSlideText,
  getSlides,
  loadPresentation,
  replaceTextInPresentation,
  replaceTextInSlide,
  savePresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: free-text replace', () => {
  it('replaceTextInPresentation rewrites a literal across every slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    // Seed something specific via the class API.
    const cls = await Presentation.load(await savePresentation(pres));
    const shape = cls.slides[0]?.shapes.find((s) => s.text.length > 0);
    if (!shape) throw new Error('expected a text shape');
    shape.setText('Hello FOO and FOO again');
    const seeded = await loadPresentation(await cls.save());

    const n = replaceTextInPresentation(seeded, 'FOO', 'BAR');
    expect(n).toBeGreaterThan(0);
    expect(getSlideText(getSlides(seeded)[0]!)).toContain('Hello BAR and BAR again');
  });

  it('replaceTextInPresentation accepts a RegExp', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const cls = await Presentation.load(await savePresentation(pres));
    const shape = cls.slides[0]?.shapes.find((s) => s.text.length > 0);
    if (!shape) throw new Error('expected a text shape');
    shape.setText('order #12345 and #67890');
    const seeded = await loadPresentation(await cls.save());

    const n = replaceTextInPresentation(seeded, /#\d+/, 'REDACTED');
    expect(n).toBeGreaterThan(0);
    expect(getSlideText(getSlides(seeded)[0]!)).toContain('REDACTED');
    expect(getSlideText(getSlides(seeded)[0]!)).not.toContain('#12345');
  });

  it('replaceTextInSlide is slide-scoped', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const cls = await Presentation.load(await savePresentation(pres));
    cls.slides[0]?.shapes.find((s) => s.text.length > 0)?.setText('slide-one says A');
    cls.slides[1]?.shapes.find((s) => s.text.length > 0)?.setText('slide-two says A');
    const seeded = await loadPresentation(await cls.save());

    const slides = getSlides(seeded);
    const n = replaceTextInSlide(slides[0]!, 'A', 'BB');
    expect(n).toBeGreaterThan(0);
    expect(getSlideText(slides[0]!)).toContain('BB');
    expect(getSlideText(slides[1]!)).toContain('says A'); // not touched

    // Sanity that getSlideShapes works with the loaded data.
    expect(getSlideShapes(slides[0]!).length).toBeGreaterThan(0);
  });
});
