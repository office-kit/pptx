// sortSlides — reorder slides by a custom comparator.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  Presentation,
  getSlideText,
  getSlides,
  loadPresentation,
  savePresentation,
  sortSlides,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: sortSlides', () => {
  it('reorders slides per the comparator', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    // Seed distinct text on each slide.
    const cls = await Presentation.load(await savePresentation(pres));
    cls.slides[0]?.shapes.find((s) => s.text.length > 0)?.setText('B-Second');
    cls.slides[1]?.shapes.find((s) => s.text.length > 0)?.setText('A-First');
    const seeded = await loadPresentation(await cls.save());

    sortSlides(seeded, (a, b) => getSlideText(a).localeCompare(getSlideText(b)));

    const ordered = getSlides(seeded).map((s) => getSlideText(s));
    expect(ordered[0]).toContain('A-First');
    expect(ordered[1]).toContain('B-Second');

    // Persistence across save → reload.
    const reloaded = await Presentation.load(await savePresentation(seeded));
    expect(reloaded.slides[0]?.text).toContain('A-First');
    expect(reloaded.slides[1]?.text).toContain('B-Second');
  });

  it('is a no-op when the comparator preserves the order', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const before = getSlides(pres).map((s) => getSlideText(s));
    sortSlides(pres, () => 0);
    const after = getSlides(pres).map((s) => getSlideText(s));
    expect(after).toEqual(before);
  });
});
