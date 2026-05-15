// findSlideByText / findSlidesByText.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  Presentation,
  findSlideByText,
  findSlidesByText,
  getSlideIndex,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findSlideByText / findSlidesByText', () => {
  it('finds the first slide containing the needle', async () => {
    // Seed deck text via the class API so we have something predictable.
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const cls = await Presentation.load(await savePresentation(pres));
    cls.slides[0]?.shapes.find((s) => s.text.length > 0)?.setText('Brand: Acme');
    cls.slides[1]?.shapes.find((s) => s.text.length > 0)?.setText('Brand: Acme again');
    const seeded = await loadPresentation(await cls.save());

    const first = findSlideByText(seeded, 'Acme');
    expect(first).not.toBeNull();
    expect(getSlideIndex(seeded, first!)).toBe(0);

    const all = findSlidesByText(seeded, 'Acme');
    expect(all).toHaveLength(2);
  });

  it('accepts a RegExp', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const cls = await Presentation.load(await savePresentation(pres));
    cls.slides[0]?.shapes.find((s) => s.text.length > 0)?.setText('Order #12345');
    const seeded = await loadPresentation(await cls.save());

    const found = findSlideByText(seeded, /#\d{5}/);
    expect(found).not.toBeNull();
  });

  it('returns null / empty when no slide matches', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(findSlideByText(pres, 'NO_SUCH_TEXT_xyz123')).toBeNull();
    expect(findSlidesByText(pres, 'NO_SUCH_TEXT_xyz123')).toEqual([]);
  });
});
