// Speaker-notes authoring (notesSlide part).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { Presentation } from '../src/api/index.ts';
import { _internalPackageOf } from '../src/api/presentation.ts';
import { partName } from '../src/internal/opc/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('L3: Slide.notes / setNotes', () => {
  it('returns null when the slide has no notesSlide', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    expect(pres.slides[0]?.notes).toBeNull();
  });

  it('creates a notesSlide part on first setNotes', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    const slide = pres.slides[0];
    if (!slide) throw new Error('expected slide');
    slide.setNotes('Speaker note line 1\nLine 2');

    const pkg = _internalPackageOf(pres);
    // Notes part exists.
    const notesPart = pkg.getPart(partName('/ppt/notesSlides/notesSlide1.xml'));
    expect(notesPart).not.toBeNull();
    // Content type override present.
    expect(notesPart?.contentType).toContain('notesSlide+xml');
    // Slide reads back the notes.
    expect(slide.notes).toBe('Speaker note line 1\nLine 2');
  });

  it('updates an existing notesSlide on subsequent setNotes', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    const slide = pres.slides[0];
    if (!slide) throw new Error('expected slide');
    slide.setNotes('first');
    slide.setNotes('second');
    expect(slide.notes).toBe('second');

    const pkg = _internalPackageOf(pres);
    // Only one notes part should exist.
    const noteParts = pkg.parts.filter((p) => p.name.startsWith('/ppt/notesSlides/notesSlide'));
    expect(noteParts.length).toBe(1);
  });

  it('round-trips through save / load', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    pres.slides[0]?.setNotes('Round-tripped notes');
    pres.slides[1]?.setNotes('Second-slide notes');

    const reloaded = await Presentation.load(await pres.save());
    expect(reloaded.slides[0]?.notes).toBe('Round-tripped notes');
    expect(reloaded.slides[1]?.notes).toBe('Second-slide notes');
  });

  // Schema validation skipped: ECMA-376's transitional pml.xsd doesn't
  // declare `<notesSlide>` as a global root element (only as a complexType),
  // so xmllint refuses to validate it standalone. PowerPoint reads it fine.
  // Schema validation for notesSlide would need the strict schema variant
  // or a wrapper element — out of scope for this iteration.
});
