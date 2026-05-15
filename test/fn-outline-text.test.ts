// getOutlineText — title + body text per slide, concatenated.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayout,
  getOutlineText,
  loadPresentation,
  setSlideBody,
  setSlideTitle,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getOutlineText', () => {
  it('joins title + body for each slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    setSlideTitle(addSlide(pres, { layout }), 'Roadmap');
    const slide = addSlide(pres, { layout });
    setSlideTitle(slide, 'Status');
    setSlideBody(slide, 'on track');

    const text = getOutlineText(pres);
    expect(text).toContain('Roadmap');
    expect(text).toContain('Status');
    expect(text).toContain('on track');
    // Slides separated by blank-line by default.
    expect(text.includes('\n\n')).toBe(true);
  });

  it('respects a custom separator', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    setSlideTitle(addSlide(pres, { layout }), 'A');
    setSlideTitle(addSlide(pres, { layout }), 'B');
    expect(getOutlineText(pres, ' | ')).toContain(' | ');
  });
});
