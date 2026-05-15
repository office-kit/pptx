// Free-function slide-size API.
//
// PowerPoint stores the slide canvas as `<p:sldSz cx="..." cy="..."/>`
// on `presentation.xml`. We expose it as EMU width/height plus an
// optional aspect-ratio hint, with presets for the two common ratios.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  SLIDE_SIZE_16_9,
  SLIDE_SIZE_4_3,
  emu,
  getSlideSize,
  loadPresentation,
  savePresentation,
  setSlideSize,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: slide size', () => {
  it('getSlideSize returns the package default', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const size = getSlideSize(pres);
    expect(size).not.toBeNull();
    expect(size?.width).toBeGreaterThan(0);
    expect(size?.height).toBeGreaterThan(0);
  });

  it('setSlideSize switches a 4:3 deck to 16:9', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    setSlideSize(pres, SLIDE_SIZE_16_9);
    const after = getSlideSize(pres);
    expect(after?.width).toBe(SLIDE_SIZE_16_9.width);
    expect(after?.height).toBe(SLIDE_SIZE_16_9.height);
    expect(after?.type).toBe('screen16x9');

    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getSlideSize(reloaded)?.type).toBe('screen16x9');
  });

  it('setSlideSize accepts arbitrary EMU dimensions', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    setSlideSize(pres, {
      width: emu(10_000_000),
      height: emu(5_000_000),
      type: 'custom',
    });
    const after = getSlideSize(pres);
    expect(after?.width).toBe(10_000_000);
    expect(after?.height).toBe(5_000_000);
    expect(after?.type).toBe('custom');
  });

  it('SLIDE_SIZE_4_3 + SLIDE_SIZE_16_9 use the canonical EMU constants', () => {
    expect(SLIDE_SIZE_4_3).toEqual({ width: 9144000, height: 6858000, type: 'screen4x3' });
    expect(SLIDE_SIZE_16_9).toEqual({ width: 12192000, height: 6858000, type: 'screen16x9' });
  });
});
