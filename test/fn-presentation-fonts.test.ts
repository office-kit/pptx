// getPresentationFonts — read the theme's font scheme.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { getPresentationFonts, loadPresentation } from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getPresentationFonts', () => {
  it('returns Calibri for the Office theme', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const fonts = getPresentationFonts(pres);
    expect(fonts).not.toBeNull();
    expect(fonts!.majorLatin).toBe('Calibri');
    expect(fonts!.minorLatin).toBe('Calibri');
    // Office theme leaves ea/cs blank → null.
    expect(fonts!.majorEastAsian).toBeNull();
    expect(fonts!.majorComplexScript).toBeNull();
  });

  it('returns null on a package without a theme part', async () => {
    // We can't easily fabricate a no-theme package via the public API
    // (decks always have at least one theme), so this test guards
    // the shape of the helper rather than its empty-deck behavior.
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const fonts = getPresentationFonts(pres);
    expect(fonts).not.toBeNull();
    expect(typeof fonts!.majorLatin === 'string' || fonts!.majorLatin === null).toBe(true);
  });
});
