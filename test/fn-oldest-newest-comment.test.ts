// getOldestComment / getNewestComment — chronology pickers.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideComment,
  findSlideLayout,
  getCommentText,
  getNewestComment,
  getOldestComment,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getOldestComment / getNewestComment', () => {
  it('returns the extremes by @dt across the deck', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    addSlideComment(slide, {
      author: { name: 'A' }, text: 'first', date: new Date('2024-01-01T00:00:00Z'),
    });
    addSlideComment(slide, {
      author: { name: 'A' }, text: 'middle', date: new Date('2025-06-01T00:00:00Z'),
    });
    addSlideComment(slide, {
      author: { name: 'A' }, text: 'last', date: new Date('2026-05-15T12:00:00Z'),
    });
    expect(getCommentText(getOldestComment(pres)!)).toBe('first');
    expect(getCommentText(getNewestComment(pres)!)).toBe('last');
  });

  it('returns null when no comment has a date', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getOldestComment(pres)).toBeNull();
    expect(getNewestComment(pres)).toBeNull();
  });
});
