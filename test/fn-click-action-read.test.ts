// getShapeClickAction — read-back parity for setShapeClickAction.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getShapeClickAction,
  getSlideIndex,
  getSlideShapes,
  getSlides,
  loadPresentation,
  moveSlide,
  removeSlide,
  savePresentation,
  setShapeClickAction,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapeClickAction', () => {
  it('returns null when no click action is set', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    expect(getShapeClickAction(shape)).toBeNull();
  });

  it('round-trips a URL action', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    setShapeClickAction(shape, { kind: 'url', url: 'https://example.com/' });
    expect(getShapeClickAction(shape)).toEqual({
      kind: 'url',
      url: 'https://example.com/',
    });
  });

  it('round-trips a slide jump', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slides = getSlides(pres);
    const target = slides[1]!;
    const shape = getSlideShapes(slides[0]!)[0]!;
    setShapeClickAction(shape, { kind: 'slide', slide: target });
    const got = getShapeClickAction(shape);
    expect(got?.kind).toBe('slide');
    // Identity compares structurally — slide opaque handle is rebuilt
    // on read, so compare on a stable text proxy.
    if (got?.kind === 'slide') {
      const { getSlideText } = await import('../src/api/index.ts');
      expect(getSlideText(got.slide)).toBe(getSlideText(target));
    }
  });

  it('resolves slide jump handles across reordering and serialization', async () => {
    const bytes = await readFile(fixture('two-slides.pptx'));
    const pres = await loadPresentation(bytes);
    const [source, target] = getSlides(pres);
    const shape = getSlideShapes(source!)[0]!;
    setShapeClickAction(shape, { kind: 'slide', slide: target! });
    const action = getShapeClickAction(shape);
    if (action?.kind !== 'slide') throw new Error('Expected slide jump');
    expect(getSlideIndex(pres, action.slide)).toBe(1);
    expect(getSlideIndex(await loadPresentation(bytes), action.slide)).toBe(-1);

    moveSlide(pres, target!, 0);
    expect(getSlideIndex(pres, action.slide)).toBe(0);
    const reloaded = await loadPresentation(await savePresentation(pres));
    const reloadedAction = getShapeClickAction(getSlideShapes(getSlides(reloaded)[1]!)[0]!);
    if (reloadedAction?.kind !== 'slide') throw new Error('Expected saved slide jump');
    expect(getSlideIndex(reloaded, reloadedAction.slide)).toBe(0);
    expect(getSlideIndex(reloaded, action.slide)).toBe(-1);

    removeSlide(pres, target!);
    expect(getSlideIndex(pres, action.slide)).toBe(-1);
  });

  it('round-trips preset navigation actions', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    for (const k of ['nextSlide', 'prevSlide', 'firstSlide', 'lastSlide'] as const) {
      setShapeClickAction(shape, { kind: k });
      expect(getShapeClickAction(shape)).toEqual({ kind: k });
    }
  });
});
