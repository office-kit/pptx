// A `slidenum` field shows the slide's own number. PowerPoint recomputes it on
// open, so the `<a:t>` cached in the file is stale the moment a slide moves —
// the preview has to count instead of reading.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  addSlide,
  addSlideTextBox,
  findSlideLayout,
  inches,
  loadPresentation,
  getSlides,
  savePresentation,
  setShapeTextField,
} from '@office-kit/pptx';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));

const deckWithFields = async (count: number, cached = '') => {
  const pres = await loadPresentation(await readFile(fixturePath));
  const layout = findSlideLayout(pres, 'Blank')!;
  for (let i = 0; i < count; i += 1) {
    const slide = addSlide(pres, { layout });
    const box = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(0.6),
      text: '',
    });
    setShapeTextField(box, 'slidenum', { text: cached });
  }
  return pres;
};

// The rendered glyphs, in document order, ignoring the SVG scaffolding.
const textOf = (svg: string): string =>
  [...svg.matchAll(/<tspan[^>]*>([^<]*)<\/tspan>|<span[^>]*>([^<]*)<\/span>/g)]
    .map((m) => m[1] ?? m[2] ?? '')
    .join('');

describe('renderSlideToSvg: slide-number fields', () => {
  it('numbers each slide by its position, not by the cached text', async () => {
    const pres = await deckWithFields(3, '99');
    const slides = getSlides(pres);
    expect(slides.map((slide) => textOf(renderSlideToSvg(pres, slide)))).toEqual(['1', '2', '3']);
  });

  it('fills in a field the deck just gained and never wrote text for', async () => {
    const pres = await deckWithFields(2);
    const slides = getSlides(pres);
    expect(textOf(renderSlideToSvg(pres, slides[1]!))).toBe('2');
  });

  it('starts at the deck’s firstSlideNum', async () => {
    // No public setter writes `firstSlideNum` — PowerPoint puts it in its page
    // setup dialog and decks that use it are rare — so the attribute goes in
    // through the file, which is also how a real such deck arrives.
    const entries = unzipSync(await savePresentation(await deckWithFields(2)));
    const part = 'ppt/presentation.xml';
    entries[part] = strToU8(
      strFromU8(entries[part]!).replace('<p:presentation', '<p:presentation firstSlideNum="7"'),
    );
    const pres = await loadPresentation(zipSync(entries));

    const slides = getSlides(pres);
    expect(slides.map((slide) => textOf(renderSlideToSvg(pres, slide)))).toEqual(['7', '8']);
  });

  it('leaves every other field type showing its cached text', async () => {
    const pres = await loadPresentation(await readFile(fixturePath));
    const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Blank')! });
    const box = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(3),
      h: inches(0.6),
      text: '',
    });
    // `datetime1` has thirteen locale-dependent siblings; a preview that
    // guessed at the format would be wrong in most of them.
    setShapeTextField(box, 'datetime1', { text: '2026-09-23' });
    expect(textOf(renderSlideToSvg(pres, slide))).toBe('2026-09-23');
  });
});
