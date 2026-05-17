// getCommentsSortedByDate — chronological timeline view.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideComment,
  findSlideLayout,
  getCommentText,
  getCommentsSortedByDate,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getCommentsSortedByDate', () => {
  it('returns comments oldest-to-newest', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    addSlideComment(slide, {
      author: { name: 'A' },
      text: 'newest',
      date: new Date('2026-05-15T12:00:00Z'),
    });
    addSlideComment(slide, {
      author: { name: 'A' },
      text: 'oldest',
      date: new Date('2024-01-01T00:00:00Z'),
    });
    addSlideComment(slide, {
      author: { name: 'A' },
      text: 'middle',
      date: new Date('2025-06-01T00:00:00Z'),
    });
    const sorted = getCommentsSortedByDate(pres);
    expect(sorted.map((c) => getCommentText(c))).toEqual(['oldest', 'middle', 'newest']);
  });

  it('returns empty when no comment has a date', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getCommentsSortedByDate(pres)).toEqual([]);
  });
});
