// Free-function comments API.
//
// Verifies legacy comments (`/ppt/commentAuthors.xml` + per-slide
// `/ppt/comments/comment{N}.xml`):
//
//   - Adding the first comment to a fresh deck bootstraps both the
//     author list and the slide's comments part (with rels).
//   - Multiple authors are deduped by (name, initials).
//   - `idx` is per-author monotonic.
//   - Removing the last comment on a slide tears down its comments
//     part and the slide → comments rel.
//   - The whole thing round-trips through save → reload.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { INTERNAL_PACKAGE, SLIDE_PART_NAME } from '../src/api/_internal-symbols.ts';
import { partName } from '../src/internal/opc/index.ts';
import { REL_TYPES } from '../src/internal/presentationml/index.ts';
import {
  addSlideComment,
  getCommentAuthor,
  getCommentAuthors,
  getCommentDate,
  getCommentPosition,
  getCommentText,
  getSlideComments,
  getSlides,
  listPackageParts,
  loadPresentation,
  removeSlideComment,
  savePresentation,
  setCommentText,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const partExists = async (presBytes: Uint8Array, partPath: string): Promise<boolean> => {
  const p = await loadPresentation(presBytes);
  return listPackageParts(p).some((part) => part.name === partPath);
};

describe('fn API: comments', () => {
  it('follows imported comment and author relationships and avoids part collisions', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [first] = getSlides(pres);
    addSlideComment(first!, { author: { name: 'Imported' }, text: 'Original' });
    const pkg = pres[INTERNAL_PACKAGE];
    for (const [source, type, oldPath, newPath, target] of [
      [
        first![SLIDE_PART_NAME],
        REL_TYPES.comments,
        '/ppt/comments/comment1.xml',
        '/ppt/comments/comment2.xml',
        '../comments/comment2.xml',
      ],
      [
        partName('/ppt/presentation.xml'),
        REL_TYPES.commentAuthors,
        '/ppt/commentAuthors.xml',
        '/review/authors.xml',
        '../review/authors.xml',
      ],
    ] as const) {
      const old = pkg.getPart(partName(oldPath))!;
      pkg.addPart(partName(newPath), old.contentType, old.data);
      pkg.removePart(partName(oldPath));
      const rels = pkg.getRels(source)!;
      rels.items.find((rel) => rel.type === type)!.target = target;
      pkg.setRels(source, rels);
    }
    const loaded = await loadPresentation(await savePresentation(pres));
    const [one, two] = getSlides(loaded);
    const comment = getSlideComments(one!)[0]!;
    expect(getCommentAuthor(comment).name).toBe('Imported');
    setCommentText(comment, '編集済み');
    addSlideComment(two!, { author: { name: 'Imported' }, text: 'Second slide' });
    expect(getCommentAuthors(loaded)).toHaveLength(1);
    expect(getSlideComments(one!).map(getCommentText)).toEqual(['編集済み']);
    const saved = await loadPresentation(await savePresentation(loaded));
    const [savedOne, savedTwo] = getSlides(saved);
    expect(getSlideComments(savedTwo!).map(getCommentText)).toEqual(['Second slide']);
    removeSlideComment(getSlideComments(savedOne!)[0]!);
    expect(getSlideComments(savedOne!)).toHaveLength(0);
    expect(getSlideComments(savedTwo!).map(getCommentText)).toEqual(['Second slide']);
    expect(listPackageParts(saved).some((part) => part.name === '/review/authors.xml')).toBe(true);
  });

  it('edits text without changing comment metadata and rejects a deleted handle', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const comment = addSlideComment(slide, {
      author: { name: 'Reviewer', initials: 'RV' },
      text: 'Before',
      date: new Date('2026-01-01T00:00:00Z'),
      position: { x: 100, y: 200 },
    });
    const author = getCommentAuthor(comment);
    setCommentText(comment, '修正済み <>&\nUpdated');
    expect(getCommentText(comment)).toBe('修正済み <>&\nUpdated');
    const loaded = await loadPresentation(await savePresentation(pres));
    const saved = getSlideComments(getSlides(loaded)[0]!)[0]!;
    expect(getCommentText(saved)).toBe('修正済み <>&\nUpdated');
    expect(getCommentAuthor(saved)).toEqual(author);
    expect(getCommentDate(saved)).toBe('2026-01-01T00:00:00.000Z');
    expect(getCommentPosition(saved)).toEqual({ x: 100, y: 200 });
    removeSlideComment(comment);
    expect(() => setCommentText(comment, 'Restore')).toThrow('no longer exists');
    expect(getSlideComments(slide)).toHaveLength(0);
  });

  it('addSlideComment bootstraps authors + comments parts on a clean slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getCommentAuthors(pres)).toHaveLength(0);
    const slide = getSlides(pres)[0]!;
    expect(getSlideComments(slide)).toHaveLength(0);

    const comment = addSlideComment(slide, {
      author: { name: 'Reviewer A' },
      text: 'Looks great, ship it.',
    });
    expect(getCommentText(comment)).toBe('Looks great, ship it.');
    expect(getCommentAuthor(comment).name).toBe('Reviewer A');
    expect(getCommentAuthor(comment).initials).toBe('R');
    expect(getCommentDate(comment)).not.toBeNull();

    const bytes = await savePresentation(pres);
    expect(await partExists(bytes, '/ppt/commentAuthors.xml')).toBe(true);
    expect(await partExists(bytes, '/ppt/comments/comment1.xml')).toBe(true);
  });

  it('dedupes authors by (name, initials) and bumps idx per-author', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;

    const c1 = addSlideComment(slide, {
      author: { name: 'Alice', initials: 'AB' },
      text: 'first',
    });
    const c2 = addSlideComment(slide, {
      author: { name: 'Alice', initials: 'AB' },
      text: 'second',
    });
    const c3 = addSlideComment(slide, {
      author: { name: 'Bob' },
      text: 'third',
    });

    expect(getCommentAuthors(pres)).toHaveLength(2);
    expect(getCommentAuthor(c1).id).toBe(getCommentAuthor(c2).id);
    expect(getCommentAuthor(c3).id).not.toBe(getCommentAuthor(c1).id);

    // Author idx should be monotonic per-author.
    const comments = getSlideComments(slide);
    expect(comments.map((c) => getCommentText(c))).toEqual(['first', 'second', 'third']);

    // Round-trip preserves everything.
    const reloaded = await loadPresentation(await savePresentation(pres));
    const reloadedSlide = getSlides(reloaded)[0]!;
    expect(getSlideComments(reloadedSlide).map((c) => getCommentText(c))).toEqual([
      'first',
      'second',
      'third',
    ]);
    expect(
      getCommentAuthors(reloaded)
        .map((a) => a.name)
        .sort(),
    ).toEqual(['Alice', 'Bob']);
  });

  it('records optional position when provided', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const c = addSlideComment(slide, {
      author: { name: 'Carla' },
      text: 'check this',
      position: { x: 1234567, y: 7654321 },
      date: new Date('2026-05-15T12:00:00.000Z'),
    });
    expect(getCommentPosition(c)).toEqual({ x: 1234567, y: 7654321 });
    expect(getCommentDate(c)).toBe('2026-05-15T12:00:00.000Z');

    const reloaded = await loadPresentation(await savePresentation(pres));
    const reloadedComment = getSlideComments(getSlides(reloaded)[0]!)[0]!;
    expect(getCommentPosition(reloadedComment)).toEqual({ x: 1234567, y: 7654321 });
    expect(getCommentDate(reloadedComment)).toBe('2026-05-15T12:00:00.000Z');
  });

  it('removeSlideComment drops the part + rel when no comments remain', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const c = addSlideComment(slide, {
      author: { name: 'Dean' },
      text: 'a',
    });
    expect(getSlideComments(slide)).toHaveLength(1);
    expect(await partExists(await savePresentation(pres), '/ppt/comments/comment1.xml')).toBe(true);

    removeSlideComment(c);
    expect(getSlideComments(slide)).toHaveLength(0);
    expect(await partExists(await savePresentation(pres), '/ppt/comments/comment1.xml')).toBe(
      false,
    );
  });

  it('getAllComments totals comments across every slide', async () => {
    const { getAllComments } = await import('../src/api/index.ts');
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getAllComments(pres)).toHaveLength(0);
    const slides = getSlides(pres);
    addSlideComment(slides[0]!, { author: { name: 'A' }, text: 'one' });
    addSlideComment(slides[0]!, { author: { name: 'A' }, text: 'two' });
    addSlideComment(slides[1]!, { author: { name: 'B' }, text: 'three' });
    expect(getAllComments(pres)).toHaveLength(3);
  });

  it('keeps unrelated comments on other slides when one is removed', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slides = getSlides(pres);
    addSlideComment(slides[0]!, { author: { name: 'X' }, text: 'first slide' });
    const c2 = addSlideComment(slides[1]!, { author: { name: 'X' }, text: 'second slide' });
    expect(getSlideComments(slides[0]!)).toHaveLength(1);
    expect(getSlideComments(slides[1]!)).toHaveLength(1);

    removeSlideComment(c2);
    expect(getSlideComments(slides[0]!)).toHaveLength(1);
    expect(getSlideComments(slides[1]!)).toHaveLength(0);

    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getSlideComments(getSlides(reloaded)[0]!)).toHaveLength(1);
    expect(getSlideComments(getSlides(reloaded)[1]!)).toHaveLength(0);
  });
});
