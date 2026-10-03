import { describe, expect, it } from 'vitest';
import {
  _internalPackageOf,
  addBlankSlide,
  createPresentation,
  getSlideNotes,
  getSlideNotesLineBreaks,
  getSlides,
  loadPresentation,
  savePresentation,
  setSlideNotes,
} from '../src/api/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

describe('speaker-notes line separator semantics', () => {
  it('replaces paragraph and soft separators even when visible text is unchanged', async () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    setSlideNotes(slide, '😀\nBeta\nGamma');
    setSlideNotes(slide, '\n', { range: { start: 2, end: 3 }, newlines: 'break' });
    expect(getSlideNotes(slide)).toBe('😀\nBeta\nGamma');
    expect(getSlideNotesLineBreaks(slide)).toEqual([
      { position: 2, kind: 'break' },
      { position: 7, kind: 'paragraph' },
    ]);

    const restored = await loadPresentation(await savePresentation(pres));
    const restoredSlide = getSlides(restored)[0]!;
    expect(getSlideNotesLineBreaks(restoredSlide)).toEqual(getSlideNotesLineBreaks(slide));
    setSlideNotes(restoredSlide, '\n', { range: { start: 2, end: 3 }, newlines: 'paragraph' });
    expect(getSlideNotes(restoredSlide)).toBe('😀\nBeta\nGamma');
    expect(getSlideNotesLineBreaks(restoredSlide)).toEqual([
      { position: 2, kind: 'paragraph' },
      { position: 7, kind: 'paragraph' },
    ]);
    if (isSchemaValidationAvailable()) {
      for (const deck of [pres, restored]) {
        const notes = _internalPackageOf(deck).parts.filter((part) =>
          part.contentType.endsWith('notesSlide+xml'),
        );
        expect(notes).toHaveLength(1);
        expectSchemaValid(new TextDecoder().decode(notes[0]!.data), 'pml');
      }
    }
  });

  it('reindexes mixed separators after replacing a range spanning paragraphs', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    setSlideNotes(slide, '😀\nBeta\nGamma\nDelta');
    setSlideNotes(slide, '\n', { range: { start: 2, end: 3 }, newlines: 'break' });
    setSlideNotes(slide, 'X\nY', { range: { start: 3, end: 13 }, newlines: 'break' });
    expect(getSlideNotes(slide)).toBe('😀\nX\nY\nDelta');
    expect(getSlideNotesLineBreaks(slide)).toEqual([
      { position: 2, kind: 'break' },
      { position: 4, kind: 'break' },
      { position: 6, kind: 'paragraph' },
    ]);
  });
});
