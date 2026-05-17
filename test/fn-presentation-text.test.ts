// getPresentationText — concatenated text across every slide.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  findSlideLayout,
  getPresentationText,
  getSlideText,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getPresentationText', () => {
  it('concatenates every slide with a form-feed by default', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank')!;
    const a = addSlide(pres, { layout });
    addSlideTextBox(a, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'alpha',
    });
    const b = addSlide(pres, { layout });
    addSlideTextBox(b, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'beta',
    });

    const all = getPresentationText(pres);
    expect(all).toContain('alpha');
    expect(all).toContain('beta');
    expect(all).toContain('\f');
  });

  it('respects a custom separator', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank')!;
    addSlideTextBox(addSlide(pres, { layout }), {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'x',
    });
    addSlideTextBox(addSlide(pres, { layout }), {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'y',
    });
    expect(getPresentationText(pres, '\n---\n')).toContain('---');
  });

  it('matches manual iteration over getSlideText', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const expected = getSlides(pres)
      .map((s) => getSlideText(s))
      .join('\f');
    expect(getPresentationText(pres)).toBe(expected);
  });
});
