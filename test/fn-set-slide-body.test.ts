// setSlideBody — write into the slide's body placeholder.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayout,
  findSlidePlaceholder,
  getShapeText,
  loadPresentation,
  setSlideBody,
  setSlideTitle,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: setSlideBody', () => {
  it('writes single-line body text', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    const slide = addSlide(pres, { layout });
    setSlideTitle(slide, 'Roadmap');
    setSlideBody(slide, 'Q1: discovery');
    expect(getShapeText(findSlidePlaceholder(slide, 'body')!)).toBe('Q1: discovery');
  });

  it('splits newlines into bullets', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    const slide = addSlide(pres, { layout });
    setSlideBody(slide, 'one\ntwo\nthree');
    expect(getShapeText(findSlidePlaceholder(slide, 'body')!)).toBe('one\ntwo\nthree');
  });

  it('throws when the slide has no body placeholder', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout });
    expect(() => setSlideBody(slide, 'x')).toThrow(/body placeholder/);
  });
});
