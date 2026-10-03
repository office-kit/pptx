// Speaker-notes run formats must survive parsing, range formatting, and a
// save/load round trip just like slide text formats.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  _internalPackageOf,
  addSlide,
  findSlideLayout,
  getSlideNotes,
  getSlideNotesParagraphEndFormat,
  getSlideNotesTextFormats,
  getSlides,
  loadPresentation,
  savePresentation,
  setSlideNotes,
  setSlideNotesFormat,
  transformSlideNotesCase,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: speaker-notes run formats', () => {
  it('does not create a notes part when changing case without notes', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    addSlide(pres, { layout: findSlideLayout(pres, 'Blank')! });
    const slide = getSlides(pres).at(-1)!;
    const partCount = _internalPackageOf(pres).parts.length;
    transformSlideNotesCase(slide, 'upper');
    expect(getSlideNotes(slide)).toBeNull();
    expect(_internalPackageOf(pres).parts.length).toBe(partCount);
  });
  it('keeps UTF-16 offsets across emoji and paragraph separators', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    addSlide(pres, { layout: findSlideLayout(pres, 'Blank')! });
    const slide = getSlides(pres).at(-1)!;
    setSlideNotes(slide, 'A😀\nBeta');
    setSlideNotesFormat(slide, { bold: true }, { range: { start: 1, end: 3 } });
    setSlideNotesFormat(slide, { italic: true }, { range: { start: 4, end: 8 } });
    expect(getSlideNotesTextFormats(slide)).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          start: 1,
          end: 3,
          format: expect.objectContaining({ bold: true }),
        }),
        expect.objectContaining({
          start: 4,
          end: 8,
          format: expect.objectContaining({ italic: true }),
        }),
      ]),
    );
    const before = await savePresentation(pres);
    const formatsBefore = getSlideNotesTextFormats(slide);
    expect(() =>
      setSlideNotesFormat(slide, { bold: false }, { range: { start: 2, end: 3 } }),
    ).toThrow(RangeError);
    expect(getSlideNotes(slide)).toBe('A😀\nBeta');
    expect(getSlideNotesTextFormats(slide)).toEqual(formatsBefore);
    const restored = getSlides(await loadPresentation(before)).at(-1)!;
    expect(getSlideNotes(restored)).toBe('A😀\nBeta');
    expect(getSlideNotesTextFormats(restored)).toEqual(getSlideNotesTextFormats(slide));
  });

  it('creates an empty notes part for paragraph-end typing format and round-trips it', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    addSlide(pres, { layout: findSlideLayout(pres, 'Blank')! });
    const slide = getSlides(pres).at(-1)!;
    const partCount = _internalPackageOf(pres).parts.length;
    setSlideNotesFormat(slide, { bold: true }, { paragraphEnd: 0 });
    expect(getSlideNotes(slide)).toBe('');
    expect(_internalPackageOf(pres).parts.length).toBeGreaterThan(partCount);
    expect(getSlideNotesParagraphEndFormat(slide, 0)).toEqual(
      expect.objectContaining({ bold: true }),
    );
    setSlideNotesFormat(slide, { bold: true, italic: true }, { paragraphEnd: 0 });
    expect(getSlideNotesParagraphEndFormat(slide, 0)).toEqual(
      expect.objectContaining({ bold: true, italic: true }),
    );
    const restored = getSlides(await loadPresentation(await savePresentation(pres))).at(-1)!;
    expect(getSlideNotesParagraphEndFormat(restored, 0)).toEqual(
      expect.objectContaining({ bold: true, italic: true }),
    );
  });

  it('validates paragraph-end indexes before changing an absent notes slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    addSlide(pres, { layout: findSlideLayout(pres, 'Blank')! });
    const slide = getSlides(pres).at(-1)!;
    const partCount = _internalPackageOf(pres).parts.length;
    expect(() => setSlideNotesFormat(slide, { bold: true }, { paragraphEnd: 1 })).toThrow(
      RangeError,
    );
    expect(getSlideNotes(slide)).toBeNull();
    expect(_internalPackageOf(pres).parts.length).toBe(partCount);
    expect(() => setSlideNotesFormat(slide, { italic: true }, { paragraphEnd: -1 })).toThrow(
      RangeError,
    );
    expect(getSlideNotes(slide)).toBeNull();
    expect(_internalPackageOf(pres).parts.length).toBe(partCount);
  });

  it('reads mixed runs and preserves them across range formatting and save/load', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout });
    const slide = getSlides(pres).at(-1)!;
    setSlideNotes(slide, 'Alpha Beta');
    const part = _internalPackageOf(pres).parts.find(
      (candidate) => candidate.name === '/ppt/notesSlides/notesSlide1.xml',
    )!;
    const xml = new TextDecoder().decode(part.data);
    part.data = new TextEncoder().encode(
      xml.replace(
        /<a:r><a:rPr[^>]*\/><a:t>Alpha Beta<\/a:t><\/a:r>/,
        '<a:r><a:rPr b="1"/><a:t>Alpha </a:t></a:r><a:r><a:rPr i="1"/><a:t>Beta</a:t></a:r>',
      ),
    );

    expect(getSlideNotes(slide)).toBe('Alpha Beta');
    expect(
      getSlideNotesTextFormats(slide).map(({ start, end, format }) => ({ start, end, format })),
    ).toEqual([
      { start: 0, end: 6, format: expect.objectContaining({ bold: true }) },
      { start: 6, end: 10, format: expect.objectContaining({ italic: true }) },
    ]);

    transformSlideNotesCase(slide, 'lower', { range: { start: 0, end: 5 } });
    expect(getSlideNotes(slide)).toBe('alpha Beta');
    expect(getSlideNotesTextFormats(slide)).toEqual([
      expect.objectContaining({
        start: 0,
        end: 6,
        format: expect.objectContaining({ bold: true }),
      }),
      expect.objectContaining({
        start: 6,
        end: 10,
        format: expect.objectContaining({ italic: true }),
      }),
    ]);
    transformSlideNotesCase(slide, 'upper', { range: { start: 0, end: 5 } });

    setSlideNotesFormat(slide, { underline: 'sng' }, { range: { start: 0, end: 3 } });
    const roundTripped = await loadPresentation(await savePresentation(pres));
    const imported = getSlides(roundTripped).at(-1)!;
    expect(getSlideNotes(imported)).toBe('ALPHA Beta');
    expect(getSlideNotesTextFormats(imported)).toEqual([
      expect.objectContaining({
        start: 0,
        end: 3,
        format: expect.objectContaining({ bold: true, underline: true }),
      }),
      expect.objectContaining({
        start: 3,
        end: 6,
        format: expect.objectContaining({ bold: true }),
      }),
      expect.objectContaining({
        start: 6,
        end: 10,
        format: expect.objectContaining({ italic: true }),
      }),
    ]);
  });
});
