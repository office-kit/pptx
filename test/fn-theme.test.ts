// `getPresentationTheme` — read the package's first color scheme.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { getPresentationTheme, loadPresentation } from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getPresentationTheme', () => {
  it('returns the Office Theme color scheme on the python-pptx default', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const theme = getPresentationTheme(pres);
    expect(theme).not.toBeNull();
    expect(theme!.name.length).toBeGreaterThan(0);

    // Each accent slot is a `#RRGGBB` literal.
    for (const slot of [
      theme!.accent1,
      theme!.accent2,
      theme!.accent3,
      theme!.accent4,
      theme!.accent5,
      theme!.accent6,
    ]) {
      expect(slot).toMatch(/^#[0-9A-F]{6}$/);
    }
    // dk1/lt1 are sysClr slots; the helper flattens them to lastClr.
    expect(theme!.dark1).toMatch(/^#[0-9A-F]{6}$/);
    expect(theme!.light1).toMatch(/^#[0-9A-F]{6}$/);
  });
});
