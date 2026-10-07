// The notes and handout masters: created with PowerPoint's default structure
// the first time they are edited, their placeholder checkboxes, the shared
// notes-page orientation and the handout's slides per page. Every result is
// saved, reloaded and validated part by part against the ECMA-376 schemas.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  createPresentation,
  getHandoutMasterPlaceholders,
  getHandoutSlidesPerPage,
  getNotesMasterPlaceholders,
  getNotesPageSize,
  getSlideNotes,
  getSlides,
  listPackageParts,
  loadPresentation,
  type PresentationData,
  readPackagePart,
  setHandoutMasterPlaceholderIncluded,
  setHandoutSlidesPerPage,
  setNotesMasterPlaceholderIncluded,
  setNotesPageOrientation,
  setSlideNotes,
} from '../src/api/index.ts';
import { expectPackageValid } from './lib/expect-package-valid.ts';

const fixture = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));
const template = async (): Promise<PresentationData> => loadPresentation(await readFile(fixture));
const decode = (bytes: Uint8Array | null): string => new TextDecoder().decode(bytes!);
const types = (placeholders: ReadonlyArray<{ type: string | null }> | null) =>
  placeholders?.map((p) => p.type) ?? null;

describe('notes master', () => {
  it('is absent until edited, then PowerPoint’s default', async () => {
    const pres = createPresentation();
    expect(getNotesMasterPlaceholders(pres)).toBeNull();

    setNotesMasterPlaceholderIncluded(pres, 'hdr', true);

    const placeholders = getNotesMasterPlaceholders(pres)!;
    expect(types(placeholders)).toEqual(['hdr', 'dt', 'sldImg', 'body', 'ftr', 'sldNum']);
    // A 16:9 slide fills the 6 in width of the slide image box.
    expect(placeholders[2]!.bounds).toEqual({ x: 685800, y: 1143000, w: 5486400, h: 3086100 });
    expect(placeholders[3]!.bounds).toEqual({ x: 685800, y: 4400550, w: 5486400, h: 3600450 });
    const presentation = decode(readPackagePart(pres, '/ppt/presentation.xml'));
    expect(presentation).toMatch(
      /<\/p:sldMasterIdLst><p:notesMasterIdLst><p:notesMasterId r:id="rId\d+"\/>/,
    );
    // The new master has a theme of its own.
    expect(
      decode(readPackagePart(pres, '/ppt/notesMasters/_rels/notesMaster1.xml.rels')),
    ).toContain('Target="../theme/theme2.xml"');
    await expectPackageValid(pres);
  });

  it('a 4:3 deck gets a 4:3 slide image', async () => {
    const pres = createPresentation({ size: '4:3' });
    setNotesMasterPlaceholderIncluded(pres, 'body', true);
    expect(getNotesMasterPlaceholders(pres)![2]!.bounds).toEqual({
      x: 1143000,
      y: 685800,
      w: 4572000,
      h: 3429000,
    });
  });

  it('removes and restores placeholders in their place', async () => {
    const pres = await template();
    setNotesMasterPlaceholderIncluded(pres, 'sldImg', false);
    setNotesMasterPlaceholderIncluded(pres, 'hdr', false);
    expect(types(getNotesMasterPlaceholders(pres))).toEqual(['dt', 'body', 'ftr', 'sldNum']);
    setNotesMasterPlaceholderIncluded(pres, 'sldImg', true);
    setNotesMasterPlaceholderIncluded(pres, 'hdr', true);
    expect(types(getNotesMasterPlaceholders(pres))).toEqual([
      'hdr',
      'dt',
      'sldImg',
      'body',
      'ftr',
      'sldNum',
    ]);
    await expectPackageValid(pres);
  });

  it('links notes slides made before the master to it', async () => {
    const pres = await template();
    const slide = addBlankSlide(pres);
    setSlideNotes(slide, 'Speak slowly');
    setNotesMasterPlaceholderIncluded(pres, 'ftr', false);
    const notes = listPackageParts(pres).find((p) => p.name.startsWith('/ppt/notesSlides/_rels/'))!;
    expect(decode(readPackagePart(pres, notes.name))).toContain('../notesMasters/notesMaster1.xml');
    const reloaded = await expectPackageValid(pres);
    expect(getSlideNotes(getSlides(reloaded)[0]!)).toBe('Speak slowly');
  });

  it('rejects an unknown placeholder type', async () => {
    const pres = await template();
    expect(() => setNotesMasterPlaceholderIncluded(pres, 'title' as 'hdr', false)).toThrow(
      /unknown placeholder type/,
    );
  });
});

describe('handout master', () => {
  it('is created with the four corner placeholders', async () => {
    const pres = await template();
    expect(getHandoutMasterPlaceholders(pres)).toBeNull();
    setHandoutMasterPlaceholderIncluded(pres, 'ftr', false);
    expect(types(getHandoutMasterPlaceholders(pres))).toEqual(['hdr', 'dt', 'sldNum']);
    setHandoutMasterPlaceholderIncluded(pres, 'ftr', true);
    const placeholders = getHandoutMasterPlaceholders(pres)!;
    expect(types(placeholders)).toEqual(['hdr', 'dt', 'ftr', 'sldNum']);
    expect(placeholders[2]!.bounds).toEqual({ x: 0, y: 8685213, w: 2971800, h: 458788 });

    // With both masters, CT_Presentation orders notes before handout.
    setNotesMasterPlaceholderIncluded(pres, 'hdr', true);
    expect(decode(readPackagePart(pres, '/ppt/presentation.xml'))).toMatch(
      /<p:notesMasterIdLst>.*<\/p:notesMasterIdLst><p:handoutMasterIdLst>/,
    );
    await expectPackageValid(pres);
  });

  it('keeps slides per page in the print settings', async () => {
    const pres = await template();
    expect(getHandoutSlidesPerPage(pres)).toBe(6);
    setHandoutSlidesPerPage(pres, 3);
    expect(getHandoutSlidesPerPage(pres)).toBe(3);
    setHandoutSlidesPerPage(pres, 'outline');
    const reloaded = await expectPackageValid(pres);
    expect(getHandoutSlidesPerPage(reloaded)).toBe('outline');
    expect(decode(readPackagePart(reloaded, '/ppt/presProps.xml'))).toContain(
      '<p:prnPr prnWhat="outline"/>',
    );
  });

  it('writes slides per page into a deck without presentation properties', async () => {
    const pres = createPresentation();
    setHandoutSlidesPerPage(pres, 9);
    expect(getHandoutSlidesPerPage(await expectPackageValid(pres))).toBe(9);
  });
});

describe('notes page orientation', () => {
  it('turns the page and moves both masters and the notes pages with it', async () => {
    const pres = await template();
    setSlideNotes(addBlankSlide(pres), 'Note');
    setNotesMasterPlaceholderIncluded(pres, 'hdr', true);
    setHandoutMasterPlaceholderIncluded(pres, 'hdr', true);
    expect(getNotesPageSize(pres)).toEqual({ width: 6858000, height: 9144000 });

    setNotesPageOrientation(pres, 'landscape');

    expect(getNotesPageSize(pres)).toEqual({ width: 9144000, height: 6858000 });
    const handout = getHandoutMasterPlaceholders(pres)!;
    // The date keeps to the right edge of the wider page.
    const dt = handout[1]!.bounds!;
    expect(dt.x + dt.w).toBeCloseTo(9144000, -4);
    const image = getNotesMasterPlaceholders(pres)![2]!.bounds!;
    expect(image.w / image.h).toBeCloseTo(4 / 3, 2);
    await expectPackageValid(pres);

    setNotesPageOrientation(pres, 'portrait');
    expect(getNotesPageSize(pres)).toEqual({ width: 6858000, height: 9144000 });
  });
});
