// removeSlideNotes — tears down the notesSlide part + slide rel.
//
// `setSlideNotes` allocates a new `/ppt/notesSlides/notesSlide{N}.xml`
// part on first use and wires the slide → notesSlide rel. The remover
// undoes exactly that: drops the part, its `.rels`, and the rel — but
// leaves the shared notesMaster alone.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getSlideNotes,
  getSlides,
  listPackageParts,
  loadPresentation,
  removeSlideNotes,
  savePresentation,
  setSlideNotes,
  type PresentationData,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const hasPart = (pres: PresentationData, part: string): boolean =>
  listPackageParts(pres).some((p) => p.name === part);

describe('fn API: removeSlideNotes', () => {
  it('drops the notes part + rel when notes were present', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    setSlideNotes(slide, 'first pass');
    expect(getSlideNotes(slide)).toBe('first pass');
    // Confirm the notesSlide part really exists in the package.
    const reloadedBefore = await loadPresentation(await savePresentation(pres));
    expect(hasPart(reloadedBefore, '/ppt/notesSlides/notesSlide1.xml')).toBe(true);

    removeSlideNotes(slide);
    expect(getSlideNotes(slide)).toBeNull();

    const reloadedAfter = await loadPresentation(await savePresentation(pres));
    expect(hasPart(reloadedAfter, '/ppt/notesSlides/notesSlide1.xml')).toBe(false);
    expect(hasPart(reloadedAfter, '/ppt/notesSlides/_rels/notesSlide1.xml.rels')).toBe(false);
  });

  it('is a no-op when the slide has no notes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[1]!;
    expect(getSlideNotes(slide)).toBeNull();
    // Should not throw, should not affect anything else.
    expect(() => removeSlideNotes(slide)).not.toThrow();
    expect(getSlideNotes(slide)).toBeNull();
  });

  it('removes only the targeted slide\'s notes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [a, b] = getSlides(pres);
    setSlideNotes(a!, 'A');
    setSlideNotes(b!, 'B');
    expect(getSlideNotes(a!)).toBe('A');
    expect(getSlideNotes(b!)).toBe('B');

    removeSlideNotes(a!);
    expect(getSlideNotes(a!)).toBeNull();
    expect(getSlideNotes(b!)).toBe('B');

    const reloaded = await loadPresentation(await savePresentation(pres));
    const [a2, b2] = getSlides(reloaded);
    expect(getSlideNotes(a2!)).toBeNull();
    expect(getSlideNotes(b2!)).toBe('B');
  });
});
