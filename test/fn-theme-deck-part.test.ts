// `getPresentationTheme` — the deck theme is the slide master's theme, not
// the first theme part by name. Google Slides exports carry the notes
// master's theme as `theme1.xml` and the slide theme as `theme2.xml`;
// reading `theme1.xml` painted a dark deck white.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import {
  getPresentationFonts,
  getPresentationTheme,
  loadPresentation,
  savePresentation,
  setPresentationFonts,
  setPresentationTheme,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const THEME_CT = 'application/vnd.openxmlformats-officedocument.theme+xml';

/**
 * `two-slides.pptx` plus a decoy theme at `/ppt/theme/theme0.xml` — first by
 * name, referenced by nothing — whose scheme is named `Decoy` with lt1
 * `#212121`. The slide master keeps pointing at `theme1.xml`.
 */
const withDecoyTheme = async (): Promise<Uint8Array> => {
  const files = unzipSync(await readFile(fixture('two-slides.pptx')));
  const original = strFromU8(files['ppt/theme/theme1.xml']!);
  const decoy = original
    .replace(/<a:clrScheme name="[^"]*"/, '<a:clrScheme name="Decoy"')
    .replace(/<a:lt1>.*?<\/a:lt1>/s, '<a:lt1><a:srgbClr val="212121"/></a:lt1>')
    .replace(/<a:latin typeface="[^"]*"/, '<a:latin typeface="Decoy Sans"');
  files['ppt/theme/theme0.xml'] = strToU8(decoy);
  const contentTypes = strFromU8(files['[Content_Types].xml']!);
  files['[Content_Types].xml'] = strToU8(
    contentTypes.replace(
      '</Types>',
      `<Override PartName="/ppt/theme/theme0.xml" ContentType="${THEME_CT}"/></Types>`,
    ),
  );
  return zipSync(files);
};

describe('fn API: deck theme follows the slide master', () => {
  it('observes theme and font changes after repeated reads', async () => {
    const pres = await loadPresentation(await withDecoyTheme());
    const original = getPresentationTheme(pres)!;
    const fonts = getPresentationFonts(pres)!;
    expect(getPresentationTheme(pres)).toEqual(original);
    expect(getPresentationFonts(pres)).toEqual(fonts);
    setPresentationTheme(pres, { accent1: '#ABCDEF' });
    setPresentationFonts(pres, { minorLatin: 'Test Sans' });
    expect(getPresentationTheme(pres)!.accent1).toBe('#ABCDEF');
    expect(getPresentationFonts(pres)!.minorLatin).toBe('Test Sans');
    expect(original.accent1).not.toBe('#ABCDEF');
    expect(fonts.minorLatin).not.toBe('Test Sans');
  });

  it('reads the master theme when an unreferenced theme sorts first by name', async () => {
    const plain = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const expected = getPresentationTheme(plain)!;
    const expectedFonts = getPresentationFonts(plain)!;

    const pres = await loadPresentation(await withDecoyTheme());
    expect(getPresentationTheme(pres)).toEqual(expected);
    expect(getPresentationTheme(pres)!.name).not.toBe('Decoy');
    expect(getPresentationFonts(pres)).toEqual(expectedFonts);
  });

  it('writes the master theme and leaves the decoy untouched', async () => {
    const pres = await loadPresentation(await withDecoyTheme());
    setPresentationTheme(pres, { accent1: '#ABCDEF' });

    const files = unzipSync(await savePresentation(pres));
    expect(strFromU8(files['ppt/theme/theme1.xml']!)).toContain('ABCDEF');
    expect(strFromU8(files['ppt/theme/theme0.xml']!)).not.toContain('ABCDEF');
  });
});
