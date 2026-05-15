// findCommentAuthorByName — author handle lookup.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideComment,
  findCommentAuthorByName,
  findSlideLayout,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findCommentAuthorByName', () => {
  it('returns the author handle for a registered name', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    addSlideComment(slide, { author: { name: 'Alice', initials: 'A' }, text: '1' });
    const found = findCommentAuthorByName(pres, 'Alice');
    expect(found).not.toBeNull();
    expect(found!.name).toBe('Alice');
    expect(typeof found!.id).toBe('number');
  });

  it('returns null when nobody matches', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    addSlideComment(slide, { author: { name: 'Alice' }, text: '1' });
    expect(findCommentAuthorByName(pres, 'NoOne')).toBeNull();
  });
});
