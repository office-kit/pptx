// findCommentsBefore — comments older than a given date.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideComment,
  findCommentsBefore,
  findSlideLayout,
  getCommentText,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findCommentsBefore', () => {
  it('returns only comments whose dt is strictly before the threshold', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    addSlideComment(slide, {
      author: { name: 'A' }, text: 'old', date: new Date('2024-01-01T00:00:00Z'),
    });
    addSlideComment(slide, {
      author: { name: 'A' }, text: 'new', date: new Date('2026-05-15T12:00:00Z'),
    });
    const hits = findCommentsBefore(pres, '2025-01-01T00:00:00Z');
    expect(hits.length).toBe(1);
    expect(getCommentText(hits[0]!)).toBe('old');
  });

  it('returns empty when nothing predates the threshold', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    addSlideComment(slide, {
      author: { name: 'A' }, text: 'future', date: new Date('2030-01-01T00:00:00Z'),
    });
    expect(findCommentsBefore(pres, new Date('2025-01-01T00:00:00Z'))).toEqual([]);
  });
});
