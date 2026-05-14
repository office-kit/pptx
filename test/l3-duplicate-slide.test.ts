// L3: duplicateSlide.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { Presentation } from '../src/api/index.ts';
import { _internalPackageOf } from '../src/api/presentation.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('L3: duplicateSlide', () => {
  it('produces a deep copy with independent text', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    const original = pres.slides[0];
    if (!original) throw new Error('expected slide');
    const dup = pres.duplicateSlide(original);

    expect(pres.slides.length).toBe(3);
    expect(dup.text).toBe(original.text);

    // Mutating the duplicate does not affect the original.
    dup.findPlaceholder('title')?.setText('Duplicated slide');
    expect(original.findPlaceholder('title')?.text).toBe('Slide 1');

    const reloaded = await Presentation.load(await pres.save());
    expect(reloaded.slides.length).toBe(3);
    expect(reloaded.slides[0]?.findPlaceholder('title')?.text).toBe('Slide 1');
    expect(reloaded.slides[2]?.findPlaceholder('title')?.text).toBe('Duplicated slide');
  });

  it('shares media parts with the source slide', async () => {
    const pres = await Presentation.load(await readFile(fixture('one-image-slide.pptx')));
    const original = pres.slides[0];
    if (!original) throw new Error('expected slide');
    pres.duplicateSlide(original);

    const pkg = _internalPackageOf(pres);
    // Still just one media part — duplicate references the same media.
    const mediaParts = pkg.parts.filter((p) => p.name.startsWith('/ppt/media/'));
    expect(mediaParts.length).toBe(1);
  });

  it('preserves layout binding on the duplicate', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    const original = pres.slides[0];
    if (!original) throw new Error('expected slide');
    const dup = pres.duplicateSlide(original);
    expect(dup.layout?.name).toBe(original.layout?.name);
  });
});
