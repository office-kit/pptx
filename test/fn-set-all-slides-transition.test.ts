// setAllSlidesTransition — apply the same transition to every slide.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getSlideTransition,
  getSlides,
  loadPresentation,
  setAllSlidesTransition,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: setAllSlidesTransition', () => {
  it('writes the requested effect onto every slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    setAllSlidesTransition(pres, { effect: 'fade' });
    for (const slide of getSlides(pres)) {
      const t = getSlideTransition(slide);
      expect(t).not.toBeNull();
      expect(t!.effect).toBe('fade');
    }
  });

  it('passes through the full TransitionOptions object', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    setAllSlidesTransition(pres, { effect: 'push', speed: 'slow' });
    for (const slide of getSlides(pres)) {
      const t = getSlideTransition(slide);
      expect(t!.effect).toBe('push');
      expect(t!.speed).toBe('slow');
    }
  });
});
