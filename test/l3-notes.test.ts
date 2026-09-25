// Speaker-notes authoring (notesSlide part).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { partName } from '../src/internal/opc/index.ts';
import {
  _internalPackageOf,
  getSlideNotes,
  getSlides,
  listPackageParts,
  loadPresentation,
  savePresentation,
  setSlideNotes,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('L3: getSlideNotes / setSlideNotes', () => {
  it('returns null when the slide has no notesSlide', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getSlideNotes(getSlides(pres)[0]!)).toBeNull();
  });

  it('creates a notesSlide part on first setSlideNotes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    setSlideNotes(slide, 'Speaker note line 1\nLine 2');

    const partNames = listPackageParts(pres).map((p) => p.name);
    expect(partNames).toContain('/ppt/notesSlides/notesSlide1.xml');
    expect(getSlideNotes(getSlides(pres)[0]!)).toBe('Speaker note line 1\nLine 2');
  });

  it('updates an existing notesSlide on subsequent setSlideNotes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    setSlideNotes(slide, 'first');
    setSlideNotes(getSlides(pres)[0]!, 'second');
    expect(getSlideNotes(getSlides(pres)[0]!)).toBe('second');

    const noteParts = listPackageParts(pres).filter((p) =>
      p.name.startsWith('/ppt/notesSlides/notesSlide'),
    );
    expect(noteParts.length).toBe(1);
  });

  it('round-trips through save / load', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slides = getSlides(pres);
    setSlideNotes(slides[0]!, 'Round-tripped notes');
    setSlideNotes(slides[1]!, 'Second-slide notes');

    const reloaded = await loadPresentation(await savePresentation(pres));
    const reloadedSlides = getSlides(reloaded);
    expect(getSlideNotes(reloadedSlides[0]!)).toBe('Round-tripped notes');
    expect(getSlideNotes(reloadedSlides[1]!)).toBe('Second-slide notes');
  });
});

it('reads and edits the notes body without changing earlier footer placeholders', async () => {
  const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
  const slide = getSlides(pres)[0]!;
  setSlideNotes(slide, 'Body text');
  const part = _internalPackageOf(pres).getPart(partName('/ppt/notesSlides/notesSlide1.xml'))!;
  const footer =
    '<p:sp><p:nvSpPr><p:cNvPr id="4" name="Footer"/><p:cNvSpPr/><p:nvPr><p:ph type="ftr"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Confidential</a:t></a:r></a:p></p:txBody></p:sp>';
  part.data = new TextEncoder().encode(
    new TextDecoder().decode(part.data).replace('<p:sp>', footer + '<p:sp>'),
  );
  expect(getSlideNotes(slide)).toBe('Body text');
  setSlideNotes(slide, '日本語のノート\nEnglish notes');
  const reloaded = await loadPresentation(await savePresentation(pres));
  expect(getSlideNotes(getSlides(reloaded)[0]!)).toBe('日本語のノート\nEnglish notes');
  const saved = _internalPackageOf(reloaded).getPart(partName('/ppt/notesSlides/notesSlide1.xml'))!;
  expect(new TextDecoder().decode(saved.data)).toContain(footer);
});

it('reads soft line breaks and field results in the notes body', async () => {
  const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
  const slide = getSlides(pres)[0]!;
  setSlideNotes(slide, 'Body text');
  const part = _internalPackageOf(pres).getPart(partName('/ppt/notesSlides/notesSlide1.xml'))!;
  part.data = new TextEncoder().encode(
    new TextDecoder()
      .decode(part.data)
      .replace(
        '</a:r>',
        '</a:r><a:br/><a:fld id="{1E094120-CA74-4B4D-9FF5-C21A65F10F82}" type="slidenum"><a:t>2</a:t></a:fld>',
      ),
  );
  expect(getSlideNotes(slide)).toBe('Body text\n2');
});
