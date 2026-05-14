// Slide-level introspection: `getSlideTransition` + `getSlideBackground`.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  clearSlideBackground,
  clearSlideTransition,
  getSlideBackground,
  getSlideTransition,
  getSlides,
  loadPresentation,
  setSlideBackground,
  setSlideTransition,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: slide introspection', () => {
  it('getSlideTransition returns null before any set, and the configured effect after', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    expect(getSlideTransition(slide)).toBeNull();

    setSlideTransition(slide, { effect: 'fade', speed: 'fast' });
    const got = getSlideTransition(slide);
    expect(got?.effect).toBe('fade');
    expect(got?.speed).toBe('fast');

    clearSlideTransition(slide);
    expect(getSlideTransition(slide)).toBeNull();
  });

  it('getSlideBackground reports inherit when no <p:bg> is set', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    expect(getSlideBackground(slide).kind).toBe('inherit');
  });

  it('getSlideBackground reads back a solid color after setSlideBackground', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    setSlideBackground(slide, '#ABCDEF');
    expect(getSlideBackground(slide)).toEqual({ kind: 'solid', color: '#ABCDEF' });

    clearSlideBackground(slide);
    expect(getSlideBackground(slide).kind).toBe('inherit');
  });
});
