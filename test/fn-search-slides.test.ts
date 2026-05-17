// searchSlides — composite text + notes search.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  findSlideLayout,
  getSlideIndex,
  getSlides,
  inches,
  loadPresentation,
  searchSlides,
  setSlideNotes,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: searchSlides', () => {
  it('matches by visible text or notes', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    // Slide 0: text match
    {
      const slide = addSlide(pres, { layout: blank });
      addSlideTextBox(slide, {
        x: inches(0),
        y: inches(0),
        w: inches(2),
        h: inches(1),
        text: 'TARGET visible',
      });
    }
    const textIdx = 0;
    // Slide 1: notes match
    {
      const slide = addSlide(pres, { layout: blank });
      setSlideNotes(slide, 'reviewer says TARGET');
    }
    const notesIdx = 1;
    // Slide 2: no match
    {
      addSlide(pres, { layout: blank });
    }
    const noMatchIdx = 2;

    const hits = searchSlides(pres, 'TARGET');
    const indices = hits.map((s) => getSlideIndex(pres, s)).sort();
    expect(indices).toContain(textIdx);
    expect(indices).toContain(notesIdx);
    expect(indices).not.toContain(noMatchIdx);
  });

  it('reports a slide once even if both text and notes match', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout: blank });
    const slide = getSlides(pres).at(-1)!;
    addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'foo',
    });
    setSlideNotes(slide, 'foo too');
    expect(searchSlides(pres, 'foo').length).toBe(1);
  });
});
