// slideHasAnimations — predicate for "does this slide animate?"

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideMedia,
  clearSlideAnimations,
  createPresentation,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  setShapeAnimation,
  slideHasAnimations,
} from '../src/api/index.ts';

const ascii = (s: string): number[] => Array.from(s, (ch) => ch.charCodeAt(0));
// Container header only — enough for signature detection, not playable.
const mp4 = (): Uint8Array =>
  new Uint8Array([0, 0, 0, 0x18, ...ascii('ftypmp42'), 0, 0, 0, 0, ...ascii('mp42isom')]);

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: slideHasAnimations', () => {
  it('returns false for slides without <p:timing>', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    for (const slide of getSlides(pres)) {
      expect(slideHasAnimations(slide)).toBe(false);
    }
  });

  it('flips true after setShapeAnimation', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0];
    if (!shape) return;
    expect(slideHasAnimations(slide)).toBe(false);
    setShapeAnimation(shape, { effect: 'fadeIn' });
    expect(slideHasAnimations(slide)).toBe(true);
  });

  // A clip's play controls live in <p:timing> too, so the element's presence
  // is not by itself an animation.
  it('stays false for a slide whose only <p:timing> holds media play controls', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    addSlideMedia(slide, {
      kind: 'video',
      data: mp4(),
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(2.25),
    });
    expect(slideHasAnimations(slide)).toBe(false);
  });

  it('reports the animation on a slide that also carries a clip, and drops back to false', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const clip = addSlideMedia(slide, {
      kind: 'video',
      data: mp4(),
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(2.25),
    });
    setShapeAnimation(clip, { effect: 'fadeIn' });
    expect(slideHasAnimations(slide)).toBe(true);
    clearSlideAnimations(slide);
    expect(slideHasAnimations(slide)).toBe(false);
  });
});
