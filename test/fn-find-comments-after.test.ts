// findCommentsAfter — comments newer than a given date.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideComment,
  findCommentsAfter,
  findSlideLayout,
  getCommentText,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findCommentsAfter', () => {
  it('returns only comments whose dt is strictly after the threshold', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    addSlideComment(slide, {
      author: { name: 'A' }, text: 'old', date: new Date('2024-01-01T00:00:00Z'),
    });
    addSlideComment(slide, {
      author: { name: 'A' }, text: 'new', date: new Date('2026-05-15T12:00:00Z'),
    });

    const hits = findCommentsAfter(pres, '2025-01-01T00:00:00Z');
    expect(hits.length).toBe(1);
    expect(getCommentText(hits[0]!)).toBe('new');
  });

  it('accepts a Date argument and filters older comments out', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    addSlideComment(slide, {
      author: { name: 'A' }, text: 'before', date: new Date('2025-01-01T00:00:00Z'),
    });
    addSlideComment(slide, {
      author: { name: 'A' }, text: 'after', date: new Date('2026-05-15T12:00:00Z'),
    });
    const hits = findCommentsAfter(pres, new Date('2026-01-01T00:00:00Z'));
    expect(hits.length).toBe(1);
    expect(getCommentText(hits[0]!)).toBe('after');
  });
});
