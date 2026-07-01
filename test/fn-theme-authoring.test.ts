// `setPresentationTheme` / `setPresentationFonts` — brand a deck's color
// scheme and typography without needing a pre-made template.

import { describe, expect, it } from 'vitest';
import {
  createPresentation,
  getPresentationFonts,
  getPresentationTheme,
  listPackageParts,
  loadPresentation,
  readPackagePart,
  savePresentation,
  setPresentationFonts,
  setPresentationTheme,
  type PresentationData,
} from '../src/api/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

const THEME_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.theme+xml';

const themePartXml = (pres: PresentationData): string => {
  const part = listPackageParts(pres).find((p) => p.contentType === THEME_CONTENT_TYPE)!;
  return new TextDecoder().decode(readPackagePart(pres, part.name)!);
};

describe('fn API: setPresentationTheme', () => {
  it('overwrites only the provided color slots, leaving the rest untouched', () => {
    const pres = createPresentation();
    const before = getPresentationTheme(pres)!;

    setPresentationTheme(pres, {
      name: 'McKinsey Navy',
      dark1: '#0B1F3A',
      accent1: '#00A9E0',
      accent2: '#FDB913',
    });

    const after = getPresentationTheme(pres)!;
    expect(after.name).toBe('McKinsey Navy');
    expect(after.dark1).toBe('#0B1F3A');
    expect(after.accent1).toBe('#00A9E0');
    expect(after.accent2).toBe('#FDB913');
    // Untouched slots keep their original value.
    expect(after.light1).toBe(before.light1);
    expect(after.accent3).toBe(before.accent3);
    expect(after.hyperlink).toBe(before.hyperlink);
  });

  it('normalizes 3-digit hex and lowercase input to uppercase #RRGGBB', () => {
    const pres = createPresentation();
    setPresentationTheme(pres, { accent1: '#0af', accent2: 'ff00aa' });
    const theme = getPresentationTheme(pres)!;
    expect(theme.accent1).toBe('#00AAFF');
    expect(theme.accent2).toBe('#FF00AA');
  });

  it('throws on an invalid color and leaves the theme unmodified', () => {
    const pres = createPresentation();
    const before = getPresentationTheme(pres)!;
    expect(() => setPresentationTheme(pres, { accent1: 'not-a-color' })).toThrow(
      /must be a #RRGGBB color/,
    );
    expect(getPresentationTheme(pres)).toEqual(before);
  });

  skipIfNoXmllint('produces a schema-valid theme part', () => {
    const pres = createPresentation();
    setPresentationTheme(pres, {
      dark1: '#0B1F3A',
      light1: '#FFFFFF',
      accent1: '#00A9E0',
    });
    expectSchemaValid(themePartXml(pres), 'dml');
  });

  it('round-trips through save/load', async () => {
    const pres = createPresentation();
    setPresentationTheme(pres, { accent1: '#123456' });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    expect(getPresentationTheme(reloaded)!.accent1).toBe('#123456');
  });
});

describe('fn API: setPresentationFonts', () => {
  it('overwrites only the provided typefaces, leaving the rest untouched', () => {
    const pres = createPresentation();
    const before = getPresentationFonts(pres)!;

    setPresentationFonts(pres, { majorLatin: 'Georgia', minorLatin: 'Verdana' });

    const after = getPresentationFonts(pres)!;
    expect(after.majorLatin).toBe('Georgia');
    expect(after.minorLatin).toBe('Verdana');
    expect(after.majorEastAsian).toBe(before.majorEastAsian);
    expect(after.minorComplexScript).toBe(before.minorComplexScript);
  });

  it('round-trips through save/load', async () => {
    const pres = createPresentation();
    setPresentationFonts(pres, { majorLatin: 'Georgia', minorLatin: 'Verdana' });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const fonts = getPresentationFonts(reloaded)!;
    expect(fonts.majorLatin).toBe('Georgia');
    expect(fonts.minorLatin).toBe('Verdana');
  });
});
