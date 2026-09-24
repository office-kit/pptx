// Free-function slide-size API.
//
// PowerPoint stores the slide canvas as `<p:sldSz cx="..." cy="..."/>`
// on `presentation.xml`. We expose it as EMU width/height plus an
// optional aspect-ratio hint, with presets for the two common ratios.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { INTERNAL_PACKAGE } from '../src/api/_internal-symbols.ts';
import { PRES_PART_NAME, decode, encode } from '../src/api/fn/_helpers.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';
import {
  createPresentation,
  SLIDE_SIZE_16_9,
  SLIDE_SIZE_4_3,
  emu,
  getSlideSize,
  getFirstSlideNumber,
  getNotesSize,
  loadPresentation,
  savePresentation,
  setSlideSize,
  addBlankSlide,
  addSlideTextBox,
  getShapeBounds,
  getSlideShapes,
  getSlides,
  getSlideXmlString,
  groupShapes,
  inches,
  setShapeTextFormat,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: slide size', () => {
  it('saves notes orientation independently, preserves paper dimensions, and validates before mutation', async () => {
    const pres = createPresentation();
    const original = getNotesSize(pres)!;
    setSlideSize(pres, SLIDE_SIZE_4_3, { notesOrientation: 'landscape' });
    expect(getNotesSize(pres)).toEqual({ width: original.height, height: original.width });
    expect(getNotesSize(await loadPresentation(await savePresentation(pres)))).toEqual(
      getNotesSize(pres),
    );
    setSlideSize(pres, SLIDE_SIZE_4_3, { notesOrientation: 'portrait' });
    expect(getNotesSize(pres)).toEqual(original);
    const part = pres[INTERNAL_PACKAGE].getPart(PRES_PART_NAME)!;
    const bytes = part.data.slice();
    expect(() =>
      setSlideSize(pres, SLIDE_SIZE_16_9, { notesOrientation: 'bad' as 'portrait' }),
    ).toThrow();
    expect(part.data).toEqual(bytes);
    part.data = encode(decode(bytes).replace(/<p:notesSz[^>]*\/>/, ''));
    expect(getNotesSize(pres)).toBeNull();
    setSlideSize(pres, SLIDE_SIZE_4_3, { notesOrientation: 'landscape' });
    expect(getNotesSize(pres)).toEqual({ width: 9144000, height: 6858000 });
    if (await isSchemaValidationAvailable()) await expectSchemaValid(decode(part.data), 'pml');
  });
  it('persists the first slide number atomically with page setup and validates its range', async () => {
    const pres = createPresentation();
    expect(getFirstSlideNumber(pres)).toBe(1);
    for (const number of [0, 10, 9999]) {
      setSlideSize(pres, SLIDE_SIZE_4_3, { firstSlideNumber: number });
      expect(getFirstSlideNumber(await loadPresentation(await savePresentation(pres)))).toBe(
        number,
      );
    }
    const bytes = pres[INTERNAL_PACKAGE].getPart(PRES_PART_NAME)!.data.slice();
    for (const firstSlideNumber of [-1, 10000, 1.5, NaN, Infinity]) {
      expect(() => setSlideSize(pres, SLIDE_SIZE_16_9, { firstSlideNumber })).toThrow();
      expect(pres[INTERNAL_PACKAGE].getPart(PRES_PART_NAME)!.data).toEqual(bytes);
    }
    setSlideSize(pres, SLIDE_SIZE_16_9);
    expect(getFirstSlideNumber(pres)).toBe(9999);
  });
  it('scales content and nested group coordinates once, keeping live slide handles current', async () => {
    const pres = createPresentation();
    setSlideSize(pres, SLIDE_SIZE_4_3);
    const slide = addBlankSlide(pres);
    const makeText = (x: number) => {
      const shape = addSlideTextBox(slide, {
        x: inches(x),
        y: inches(1),
        w: inches(2),
        h: inches(1),
        text: 'Scaled',
      });
      setShapeTextFormat(shape, { size: 20 });
      return shape;
    };
    makeText(1);
    const second = makeText(4);
    groupShapes([getSlideShapes(slide)[0]!, second]);
    const ungrouped = makeText(7);
    const group = getSlideShapes(slide)[0]!;
    const originalGroup = getShapeBounds(group)!;
    setSlideSize(
      pres,
      { width: inches(5), height: inches(5), type: 'custom' },
      { scaleContent: true },
    );
    const expected = (bounds: NonNullable<ReturnType<typeof getShapeBounds>>) => ({
      x: bounds.x / 2,
      y: bounds.y / 2 + inches(0.625),
      w: bounds.w / 2,
      h: bounds.h / 2,
    });
    expect(getShapeBounds(group)).toEqual(expected(originalGroup));
    expect(getShapeBounds(ungrouped)).toEqual({
      x: inches(3.5),
      y: inches(1.125),
      w: inches(1),
      h: inches(0.5),
    });
    const xml = getSlideXmlString(slide);
    expect(xml).toContain('sz="1000"');
    expect(xml).toContain('<a:chOff x="457200" y="457200"');
    expect(xml).toContain('<a:chExt cx="2286000" cy="457200"');
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    expect(getShapeBounds(getSlideShapes(getSlides(reloaded)[0]!)[0]!)).toEqual(
      expected(originalGroup),
    );
    if (isSchemaValidationAvailable()) expectSchemaValid(xml, 'pml');
  });
  it('getSlideSize returns the package default', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const size = getSlideSize(pres);
    expect(size).not.toBeNull();
    expect(size?.width).toBeGreaterThan(0);
    expect(size?.height).toBeGreaterThan(0);
  });

  it('prepares all scaled parts before committing and preserves unrelated parts', async () => {
    const pres = createPresentation();
    addBlankSlide(pres);
    const pkg = pres[INTERNAL_PACKAGE];
    const layout = pkg.parts.find((part) => part.contentType.endsWith('.slideLayout+xml'))!;
    const original = layout.data;
    layout.data = encode('<invalid');
    const before = pkg.parts.map((part) => part.data);
    expect(() => setSlideSize(pres, SLIDE_SIZE_4_3, { scaleContent: true })).toThrow();
    expect(pkg.parts.map((part) => part.data)).toEqual(before);
    layout.data = original;
    const unrelated = pkg.parts.filter(
      (part) =>
        !/presentationml\.(presentation\.main|slide|slideLayout|slideMaster)\+xml$/.test(
          part.contentType,
        ),
    );
    const bytes = unrelated.map((part) => part.data);
    setSlideSize(pres, SLIDE_SIZE_4_3, { scaleContent: true });
    expect(unrelated.map((part) => part.data)).toEqual(bytes);
    for (const part of pkg.parts.filter((part) =>
      /presentationml\.(slide|slideLayout|slideMaster)\+xml$/.test(part.contentType),
    )) {
      if (isSchemaValidationAvailable()) expectSchemaValid(decode(part.data), 'pml');
    }
  });

  it('setSlideSize switches a 4:3 deck to 16:9', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    setSlideSize(pres, SLIDE_SIZE_16_9);
    const after = getSlideSize(pres);
    expect(after?.width).toBe(SLIDE_SIZE_16_9.width);
    expect(after?.height).toBe(SLIDE_SIZE_16_9.height);
    expect(after?.type).toBe('screen16x9');

    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getSlideSize(reloaded)?.type).toBe('screen16x9');
  });

  it('setSlideSize accepts arbitrary EMU dimensions', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    setSlideSize(pres, {
      width: emu(10_000_000),
      height: emu(5_000_000),
      type: 'custom',
    });
    const after = getSlideSize(pres);
    expect(after?.width).toBe(10_000_000);
    expect(after?.height).toBe(5_000_000);
    expect(after?.type).toBe('custom');
  });

  it('validates dimensions before mutation and accepts both schema limits', () => {
    const pres = createPresentation();
    const part = pres[INTERNAL_PACKAGE].getPart(PRES_PART_NAME)!;
    const before = part.data;
    for (const size of [NaN, Infinity, -Infinity, 0, 914399, 51206401]) {
      for (const dimension of ['width', 'height']) {
        expect(() => setSlideSize(pres, { ...SLIDE_SIZE_4_3, [dimension]: size })).toThrow(
          RangeError,
        );
        expect(part.data).toEqual(before);
      }
    }
    setSlideSize(pres, { width: emu(914400), height: emu(51206400), type: 'custom' });
    expect(getSlideSize(pres)).toEqual({ width: 914400, height: 51206400, type: 'custom' });
    if (isSchemaValidationAvailable()) expectSchemaValid(decode(part.data), 'pml');
  });

  it('inserts an absent size before notes properties even when there are no slides', async () => {
    const pres = createPresentation();
    const part = pres[INTERNAL_PACKAGE].getPart(PRES_PART_NAME)!;
    part.data = encode(
      decode(part.data)
        .replace(/<p:sldSz[^>]*\/>/, '')
        .replace(/<p:sldIdLst\s*\/>/, ''),
    );
    expect(getSlideSize(pres)).toBeNull();
    setSlideSize(pres, { width: emu(7000000), height: emu(11000000), type: 'custom' });
    const xml = decode(part.data);
    expect(xml.indexOf('<p:sldSz')).toBeLessThan(xml.indexOf('<p:notesSz'));
    if (isSchemaValidationAvailable()) expectSchemaValid(xml, 'pml');
    expect(getSlideSize(await loadPresentation(await savePresentation(pres)))).toEqual({
      width: 7000000,
      height: 11000000,
      type: 'custom',
    });
  });

  it('SLIDE_SIZE_4_3 + SLIDE_SIZE_16_9 use the canonical EMU constants', () => {
    expect(SLIDE_SIZE_4_3).toEqual({ width: 9144000, height: 6858000, type: 'screen4x3' });
    expect(SLIDE_SIZE_16_9).toEqual({ width: 12192000, height: 6858000, type: 'screen16x9' });
  });
});
