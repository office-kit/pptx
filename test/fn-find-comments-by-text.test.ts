// findCommentsByText — content search across all comments.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideComment,
  findCommentsByText,
  findSlideLayout,
  getCommentText,
  getSlides,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findCommentsByText', () => {
  it('matches by substring', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout: blank });
    addSlide(pres, { layout: blank });
    const slides = getSlides(pres);
    addSlideComment(slides[0]!, { author: { name: 'A' }, text: 'check the API contract' });
    addSlideComment(slides[1]!, { author: { name: 'B' }, text: 'unrelated' });
    addSlideComment(slides[1]!, { author: { name: 'A' }, text: 'API rename needed' });
    const hits = findCommentsByText(pres, 'API');
    expect(hits.length).toBe(2);
    for (const c of hits) expect(getCommentText(c)).toMatch(/API/);
  });

  it('matches by RegExp', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout: blank });
    const slides = getSlides(pres);
    addSlideComment(slides[0]!, { author: { name: 'A' }, text: 'meeting on 2026-05-15' });
    expect(findCommentsByText(pres, /\d{4}-\d{2}-\d{2}/).length).toBe(1);
    expect(findCommentsByText(pres, /^xyz/).length).toBe(0);
  });
});
